// Name: Dokument umregistrieren (RerecordDocument)
// Version: 1.0.2
// Datum: 27.06.2025
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ActaNova.Domain;
using ActaNova.Domain.Classifications;
using ActaNova.Domain.TransferOperations.Rerecord;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Domain;
using Rubicon.Domain.TransferOperations;
using Rubicon.Gever.Bund.Deletion.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using Document = ActaNova.Domain.Document;

namespace Bukonf.Gever.Workflow
{
  [VersionInfo ("1.0.2", "27.06.2025", "RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu")]
  public class RerecordDocumentActivityCommandModule : GeverActivityCommandModule, ITemplateOnlyActivityCommandModule
  {
    [AttributeUsage (AttributeTargets.Class, Inherited = false)]
    private class VersionInfoAttribute : Attribute
    {
      public string Version { get; }
      public string Date { get; }
      public string Author { get; }

      public VersionInfoAttribute (string version, string date, string author)
      {
        Version = version;
        Date = date;
        Author = author;
      }
    }

    public class RerecordDocumentParameters : GeverActivityCommandParameters
    {
      [Required]
      [ExpectedExpressionType (typeof (IEnumerable<Document>))]
      public TypedParsedExpression Documents { get; set; }

      [Required]
      [ExpectedExpressionType (typeof (FileHandlingContainer))]
      public TypedParsedExpression Target { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression ShowVersion { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht leer sein.")]
      ParameterNullError,

      [De (
          "Befehlsaktivität \"{0}\": Das Dokument \"{1}\" muss entweder im aktuellen Dossier oder in jenem Dossier abgelegt sein, das den betroffenen Geschäftsvorfall enthält.")]
      DocumentNotInFileError,

      [De (
          "Befehlsaktivität \"{0}\": Das Dokument \"{1}\" muss im Geschäftsvorfall referenziert sein.")]
      DocumentNotReferencedInFileCaseError,

      [De ("Befehlsaktivität \"{0}\": Das Dokument \"{1}\" kann nicht umregistriert werden.")]
      CannotBeRecordedError,

      [De ("Befehlsaktivität \"{0}\": Die Aktivität muss schreibgeschützt sein.")]
      InstanceNotReadOnlyError,

      [De ("Befehlsaktivität \"{0}\": Das Zieldossier ist ungültig.")]
      TargetInvalidError,

      [De ("Befehlsaktivität \"{0}\": Das Zieldossier kann nicht bearbeitet werden.")]
      CannotEditTargetException,

      [De ("Befehlsaktivität \"{0}\": Der Benutzer ist nicht berechtigt, das Dokument \"{1}\" umzuregistrieren.")]
      NoRerecordRightsForDocumentException,

      [De ("Befehlsaktivität \"{0}\": Version: \"{1}\" - Datum: \"{2}\" - Autor \"{3}\"")]
      ShowVersion,
    }

    public RerecordDocumentActivityCommandModule ()
        : base ("RerecordDocument:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<RerecordDocumentParameters> (commandActivity);
      var documentList = parameters.Documents;
      var targetFile = parameters.Target;

      var documents = documentList?.Invoke<IEnumerable<Document>> (commandActivity, commandActivity.WorkItem);
      var target = targetFile?.Invoke<FileHandlingContainer> (commandActivity, commandActivity.WorkItem);

      if (documents == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(RerecordDocumentParameters.Documents)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(RerecordDocumentParameters.Documents)));
      }

      if (target == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(RerecordDocumentParameters.Target)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(RerecordDocumentParameters.Target)));
      }

      if (!target.CanEdit (true))
      {
        throw new ActivityCommandException (
                $"The target cannot be edited in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.CannotEditTargetException.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName));
      }

      var sourceFile = GetSourceFile (commandActivity);
      var sourceFileDocumentList = sourceFile.GetDocumentsRecursive().ToList();

      foreach (var document in documents)
      {
        if (!document.CanBeReRecorded())
        {
          throw new ActivityCommandException (
              $"The User has no rights to rerecord the Document '{document.DisplayName}' in activity '{commandActivity.ID}'.").WithUserMessage (
              ModuleLocalization.NoRerecordRightsForDocumentException.ToLocalizedName().FormatWith (commandActivity.DisplayName, document.DisplayName));
        }

        if (!sourceFileDocumentList.Contains (document))
        {
          throw new ActivityCommandException (
                  $"The Document '{document.DisplayName}' must be included in the File or BaseFile of the FileCase in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.DocumentNotInFileError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, document.DisplayName));
        }

        if (commandActivity.EffectiveWorkItemType == typeof (FileCase))
        {
          if (((FileCase)commandActivity.WorkItem).FileCaseContents.All (f => f.Document != document))
          {
            throw new ActivityCommandException (
                $"The Document '{document.DisplayName}' must be referenced in the FileCase in activity '{commandActivity.ID}'.").WithUserMessage (
                ModuleLocalization.DocumentNotReferencedInFileCaseError.ToLocalizedName().FormatWith (commandActivity.DisplayName, document.DisplayName));
          }
        }

        var affectedObjects = document.ToEnumerable();
        var objectsWithTimestamps = affectedObjects.ToDictionary (o => o.ID, o => (byte[])o.Timestamp);
        var operation = TransferOperationModule.Value (objectsWithTimestamps, target, null, TransferOperationTargetKindType.Object)
            .OfType<RerecordTransferOperation>().Single();

        if (operation == null)
        {
          throw new ActivityCommandException ($"{document.DisplayName} can not be rerecorded to {target.DisplayName} in activity '{commandActivity.ID}'.")
              .WithUserMessage (
                  ModuleLocalization.CannotBeRecordedError.ToLocalizedName().FormatWith (commandActivity.DisplayName, document.DisplayName));
        }

        var failingDetails = operation.GetFailingDetails();
        if (failingDetails.Any())
        {
          var failReasons = failingDetails.Select (o => o.Reason).Join (", ");
          throw new ActivityCommandException (
                  $"Rerecording the document '{document.DisplayName}' is not possible due to the following reasons: '{failReasons}' in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.CannotBeRecordedError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, document.DisplayName));
        }

        document.Rerecord (target);
        document.Folder = target.RootFolder;

        if (commandActivity.EffectiveWorkItemType == typeof (FileCase))
        {
          var fileCase = (FileCase)commandActivity.WorkItem;
          var fileCaseContentToDelete = fileCase.FileCaseContents.FirstOrDefault (f => f.Document == document);
          if (fileCaseContentToDelete.CanDelete())
            fileCaseContentToDelete?.Delete();
          var connection = FileHandlingObjectConnection.NewObject();
          connection.ReferencedFileHandlingObject = document;
          connection.ClassificationType = new ReferenceHandle<RelationClassificationType> ("!RelationClassificationType!FormerFileCaseContent").GetObject();
          connection.FileHandlingObject = document;
          fileCase.FileHandlingObjectConnections.Add (connection);
        }
      }

      return true;
    }

    private File GetSourceFile (CommandActivity commandActivity)
    {
      File file;
      if (commandActivity.EffectiveWorkItemType == typeof (FileCase))
        file = (File)((FileCase)commandActivity.WorkItem).BaseFile;
      else
      {
        file = (File)commandActivity.WorkItem;
      }

      return file;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<RerecordDocumentParameters> (commandActivity);
        ShowVersion (commandActivity, parameters);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);

        var typeValidator = new WorkItemTypeValidator (typeof (File), typeof (FileCase));
        typeValidator.ValidateWorkItemType (commandActivity);

        if (!commandActivity.InstanceReadOnly)
        {
          throw new ActivityCommandException ($"Activity must be set to instance read only in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.InstanceNotReadOnlyError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    private void ShowVersion (CommandActivity commandActivity, RerecordDocumentParameters parameters)
    {
      var showVersion = parameters.ShowVersion?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      if (showVersion == true)
      {
        var attr = GetType().GetCustomAttribute<VersionInfoAttribute>();
        if (attr != null)
        {
          throw new ActivityCommandException (
                  $"CommandActivity \"{commandActivity.DisplayName}\": Version: \"{attr.Version}\" - Date: \"{attr.Date}\" - Author \"{attr.Author}\"")
              .WithUserMessage (ModuleLocalization
                  .ShowVersion.ToLocalizedName().FormatWith (
                      commandActivity.DisplayName,
                      attr.Version,
                      attr.Date,
                      attr.Author));
        }
      }
    }
  }
}