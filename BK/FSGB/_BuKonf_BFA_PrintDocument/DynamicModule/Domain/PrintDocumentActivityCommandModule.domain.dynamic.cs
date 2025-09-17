// Name: Dokument drucken (PrintDocument)
// Version: 1.0.0
// Datum: 22.08.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using ActaNova.Domain;
using ActaNova.Domain.Printing;
using ActaNova.Domain.Workflow;
using Aspose.Pdf.Facades;
using Remotion.Globalization;
using Remotion.Security;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Security;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  public class PrintDocumentActivityCommandModule : GeverActivityCommandModule
  {
    public class PrintDocumentParameters : GeverActivityCommandParameters
    {
      [Required]
      [ExpectedExpressionType (typeof (PrinterClassificationType))]
      public TypedParsedExpression Printer { get; set; }

      [ExpectedExpressionType (typeof (IEnumerable<Document>))]
      public TypedParsedExpression Documents { get; set; }

      public bool All { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Es wurden keine PDF Inhalte zum drucken gefunden.")]
      NoPdfContentError
    }

    public PrintDocumentActivityCommandModule ()
        : base ("PrintDocument:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<PrintDocumentParameters> (commandActivity);

      var printer = parameters.Printer?.Invoke<PrinterClassificationType> (commandActivity, commandActivity.WorkItem);
      CheckParameterNotNull (commandActivity, printer, nameof(PrintDocumentParameters.Printer));

      var documents = parameters.Documents?.Invoke<IEnumerable<Document>> (commandActivity, commandActivity.WorkItem);
      if (parameters.All)
        documents = ((FileHandlingContainer)commandActivity.WorkItem).Documents.Where (d => d.IsPdfContentAvailable);

      CheckParameterNotNull (commandActivity, documents, nameof(PrintDocumentParameters.Documents));

      documents = documents.Where (d => AccessControlUtility.HasAccess (d, GeneralAccessTypes.Read));

      var documentsArray = documents as Document[] ?? documents.ToArray();
      if (!documentsArray.Any (d => d.IsPdfContentAvailable))
      {
        throw new ActivityCommandException (
                $"No PDF Content found to print in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.NoPdfContentError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName));
      }

      var pdfContents = documentsArray.Select (d => d.PdfContent);

      using (var viewer = new PdfViewer())
      {
        viewer.AutoResize = true;
        viewer.AutoRotate = true;
        viewer.PrintPageDialog = false;

        foreach (var pdfContent in pdfContents)
        {
          using (var stream = pdfContent.GetContent().CreateStream (true))
          {
            viewer.BindPdf (stream);
            var ps = new PrinterSettings { PrinterName = printer.PrinterName };

            viewer.PrintDocumentWithSettings (ps);
          }
        }

        viewer.Close();
      }

      return true;
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
        var typeValidator = new WorkItemTypeValidator (typeof (FileHandlingContainer));
        typeValidator.ValidateWorkItemType (commandActivity);

        var parameters = ActivityCommandParser.Parse<PrintDocumentParameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);

        if (parameters.All == false && parameters.Documents == null)
        {
          throw new ActivityCommandException (
                  $"Parameter '{nameof(PrintDocumentParameters.Documents)}' must not be null in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, nameof(PrintDocumentParameters.Documents)));
        }
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }
  }
}