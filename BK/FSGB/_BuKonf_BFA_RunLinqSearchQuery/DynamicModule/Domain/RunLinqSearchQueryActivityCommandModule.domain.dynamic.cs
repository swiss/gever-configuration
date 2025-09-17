// Name: Expertensuche ausführen (RunLinqSearchQuery)
// Version: 1.0.1
// Datum: 03.11.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu

using System;
using System.IO;
using System.Linq;
using ActaNova.Domain;
using ActaNova.Domain.Classifications;
using ActaNova.Domain.LinqSearch;
using ActaNova.Domain.Workflow;
using Newtonsoft.Json;
using Remotion.Globalization;
using Remotion.Utilities;
using Rubicon.Dms;
using Rubicon.Domain;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.TenantCross;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.IO;
using Rubicon.Utilities.Security;
using Rubicon.Workflow.Domain;
using Document = ActaNova.Domain.Document;
using File = ActaNova.Domain.File;

namespace Bukonf.Gever.Workflow
{
  public class RunLinqSearchQueryActivityCommandModule : GeverActivityCommandModule, ITemplateOnlyActivityCommandModule
  {
    public class ModuleParameters : GeverActivityCommandParameters
    {
      [Required]
      [ExpectedExpressionType (typeof (LinqSearchQuery))]
      public TypedParsedExpression Query { get; set; }

      [Required]
      [ExpectedExpressionType (typeof (FileHandlingContainer))]
      public TypedParsedExpression SaveLocation { get; set; }

      public string QueryParameter { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Abfragen im Status \"Entwurf\" können nicht ausgeführt werden.")]
      CannnotExecuteDraft,

      [De ("Befehlsaktivität \"{0}\": Mandantenübergreifende Abfragen können nicht ausgeführt werden.")]
      CannnotExecuteTenantCross,

      [De ("Befehlsaktivität \"{0}\": Nicht alle verpflichtenden Parameter haben einen Wert.")]
      RequiredParametersNotSet,

      [De ("Die Aktivität muss schreibgeschützt in der Prozessinstanz sein.")]
      InstanceNotReadOnlyError,

      [De ("Befehlsaktivität \"{0}\": Syntaxfehler. Geben Sie den Parameter mit korrekter Abfragesyntax an. Erwarteter Ergebnistyp: \"{1}\"")]
      ExpressionParseError,

      [De ("Befehlsaktivität \"{0}\": Der angegebene Speicherort entspricht nicht den Anforderungen.")]
      InvalidSaveLocationError,

      [De ("Befehlsaktivität \"{0}\": Der Benutzer ist nicht berechtigt einen Excel Export durchzuführen.")]
      CannotExportExcel,

      [De("Befehlsaktivität \"{0}\": Die Expertensuche \"{1}\" muss eine Excel-Vorlage definieren.")]
      MissingExcelTemplateError,
    }

    public RunLinqSearchQueryActivityCommandModule ()
        : base ("RunLinqSearchQuery:ActivityCommandClassificationType")
    {
    }

    private const string c_resultDocumentFlowVariableName = "ResultDocument";

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<ModuleParameters> (commandActivity);

      var searchQuery = parameters.Query?.Invoke<LinqSearchQuery> (commandActivity, commandActivity.WorkItem);
      CheckParameterNotNull (commandActivity, searchQuery, nameof(ModuleParameters.Query));

      var saveLocation = parameters.SaveLocation?.Invoke<FileHandlingContainer> (commandActivity, commandActivity.WorkItem);
      CheckParameterNotNull (commandActivity, searchQuery, nameof(ModuleParameters.SaveLocation));

      if (saveLocation != commandActivity.WorkItem
          && !(commandActivity.WorkItem is FileCase fc && saveLocation == fc.BaseFile)
          && !(commandActivity.WorkItem is FileHandlingContainer fhc && saveLocation == fhc.ParentFileContentHierarchyObject)
          && !(commandActivity.WorkItem is File file && file.SubFiles.Contains (saveLocation))
          || !saveLocation.CanEdit())
      {
        throw new ActivityCommandException (
                $"Save location does not meet the requirements in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.InvalidSaveLocationError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName));
      }

      if (((ITenantCrossLinqSearchQueryMixin)searchQuery).AllowTenantCrossSearch)
      {
        throw new ActivityCommandException (
                $"TenantCross LinqSearchQuery not supported in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.CannnotExecuteTenantCross.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName));
      }

      if (searchQuery.QueryState != LinqSearchQueryStateType.Executeable)
      {
        throw new ActivityCommandException (
                $"LinqSearchQuery must no be in draft state in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.CannnotExecuteDraft.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName));
      }

      dynamic queryParameter = null;

      if (parameters.QueryParameter != null)
      {
        queryParameter = JsonConvert.DeserializeObject (parameters.QueryParameter);
      }

      if (searchQuery.Parameters.Any() && queryParameter != null)
      {
        searchQuery.EnsureParameterCacheInvalidated();

        var staticParameters = searchQuery.Parameters
            .Where (p => !p.IsCalculated)
            .ToDictionary (p => p.ParameterName);

        foreach (var property in queryParameter)
        {
          var parameter = staticParameters[(string)property.identifier];

          var actualValue = GetValueFromExpression (commandActivity, (string)property.value);

          var value = Convert.ChangeType (actualValue, parameter.ParameterType);
          parameter.ParameterValue = value;
        }
      }

      var requiredParametersSet = searchQuery.Parameters.Where (p => p.IsRequired && !p.IsCalculated).All (p => p.ParameterValue != null);
      if (!requiredParametersSet)
      {
        throw new ActivityCommandException (
                $"Not all required parameters have a value in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.RequiredParametersNotSet.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName));
      }

      if (!AccessControlUtility.HasAccess (searchQuery, LinqSearchQuery.AccessTypes.Execute) || !searchQuery.CanUserExportExcel())
      {
        throw new ActivityCommandException (
                $"User can not export excel in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.CannotExportExcel.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName));
      }

      var resultDocument = ExecuteToExcel (searchQuery, saveLocation, commandActivity);

      var errorMessageFlowVariable = commandActivity.GetOrCreateFlowVariable (c_resultDocumentFlowVariableName);
      errorMessageFlowVariable.SetValue (resultDocument);

      return true;
    }

    private object GetValueFromExpression (CommandActivity commandActivity, string expressionString)
    {
      var workItemType = commandActivity.WorkItem?.GetPublicDomainObjectType() ?? commandActivity.EffectiveWorkItemType;
      var formattedExpressionString = $"(activity, workItem) => {expressionString}";

      var parsedExpression = ExpressionConfiguration.Default.Parse (formattedExpressionString, typeof (CommandActivity), workItemType)
          .WithExpectedReturnType (typeof (object));

      var parseException = parsedExpression.Exception;
      if (!parsedExpression.IsValid || parseException != null)
      {
        throw new ActivityCommandException ($"Error while parsing expression '{expressionString}' of activity '{commandActivity.ID}'.", parseException)
            .WithUserMessage (ModuleLocalization.ExpressionParseError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(Object)));
      }

      var actualObject = parsedExpression?.Invoke<object> (commandActivity, commandActivity.WorkItem);
      return actualObject;
    }

    private Document ExecuteToExcel (LinqSearchQuery linqSearchQuery, FileHandlingContainer saveLocation, CommandActivity commandActivity)
    {
      ArgumentUtility.CheckNotNull (nameof(linqSearchQuery), linqSearchQuery);

      var exportTitle = linqSearchQuery.Name;

      Stream templateStream = null;
      var excelTemplate = linqSearchQuery.ExcelTemplate;
      if (excelTemplate != null)
      {
        using (var handle = excelTemplate.PrimaryContent.GetContent())
        {
          using (var stream = handle.CreateStream())
          {
            templateStream = new MemoryStream();
            stream.CopyTo (templateStream);
          }
        }
      }
      else
      {
        throw new ActivityCommandException ($"Linq search query '{exportTitle}' must provide an ExcelTemplate.")
            .WithUserMessage (ModuleLocalization.MissingExcelTemplateError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, exportTitle));
      }

      const string mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      var extension = excelTemplate.PrimaryContent.GetExtension();
      var documentName = exportTitle + "_" + DateTime.Now.ToString ("yyyy-MM-dd_HH-mm-ss");

      var excelExportDocument = CreateExport (linqSearchQuery, templateStream);

      var data = excelExportDocument.Save();

      var newDocument = Document.NewObject (saveLocation);
      newDocument.PrimaryContent.Name = documentName;
      newDocument.Type = DocumentClassificationType.WellKnown.CommonDocumentType.GetObject();

      using (var handle = new SimpleStreamAccessHandle (s => new MemoryStream (data)))
      {
        newDocument.PrimaryContent.SetContent (handle, extension.ToString(), mimeType);
      }

      return newDocument;
    }

    private ISaveableDocument CreateExport (LinqSearchQuery linqSearchQuery, Stream templateStream)
    {
      var exportTitle = linqSearchQuery.SheetOrTableName ?? linqSearchQuery.Name;
      var exportDelegate = linqSearchQuery.CreateExcelExportDelegate (exportTitle, templateStream);
      return exportDelegate();
    }

    private void CheckParameterNotNull (CommandActivity commandActivity, object invokedObject, string parameterName)
    {
      if (invokedObject == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{parameterName}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.ParameterNullError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, parameterName));
      }
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        if (!commandActivity.InstanceReadOnly)
        {
          throw new ActivityCommandException ($"Activity must be set to instance read only'{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.InstanceNotReadOnlyError.ToLocalizedName());
        }

        var parameters = ActivityCommandParser.Parse<ModuleParameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }
  }
}