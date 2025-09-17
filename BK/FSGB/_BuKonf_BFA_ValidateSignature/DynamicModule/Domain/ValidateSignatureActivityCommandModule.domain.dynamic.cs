// Name: Signatur validieren (validateSignature)
// Version: 1.0.1
// Datum: 15.01.2024
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.IO;
using System.Text.RegularExpressions;
using ActaNova.Domain;
using ActaNova.Domain.Security;
using ActaNova.Domain.Workflow;
using Aspose.Email.Mime;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Utilities.CommandLineTools;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using static System.IO.Directory;
using Document = ActaNova.Domain.Document;
using File = ActaNova.Domain.File;

namespace Bukonf.Gever.Workflow
{
  [LocalizationEnum]
  public enum ModuleLocalization
  {
    [De ("Befehlsaktivität \"{0}\": Das Aktivitätsobjekt muss vom Typ Dossier, Eingang oder Geschäftsvorfall sein.")]
    InvalidWorkItemTypeError,

    [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
    ParameterNullError,

    [De ("Befehlsaktivität \"{0}\": Das Dokument \"{1}\" hat keinen aktiven Inhalt.")]
    MissingContentError,

    [De ("Befehlsaktivität \"{0}\": Das temporäre Verzeichnis konnte nicht angelegt werden.")]
    FolderSetupError,

    [De ("Befehlsaktivität \"{0}\": Das temporäre Verzeichnis konnte nicht gelöscht werden.")]
    FolderDeletionError,

    [De ("Befehlsaktivität \"{0}\": Beim Ausführen der Signaturvalidierung ist ein Fehler aufgetreten.")]
    SignatureValidationError,

    [De ("Befehlsaktivität \"{0}\": Beim Herunterladen des PDF ist ein Fehler aufgetreten.")]
    DownloadPdfError,

    [De ("Befehlsaktivität \"{0}\": Beim Hochladen des Validierungsreport ist ein Fehler aufgetreten.")]
    UploadReportError,

    [De ("Befehlsaktivität \"{0}\": Der aktuelle Benutzer hat nicht genügend Rechte um das Dokument \"{1}\" zu lesen.")]
    ReadAccessDeniedError,
  }

  public class ValidateSignatureActivityCommandModule : GeverActivityCommandModule
  {
    public class ModuleParameters : GeverActivityCommandParameters
    {
      [Required]
      [ExpectedExpressionType (typeof (Document))]
      public TypedParsedExpression Document { get; set; }
    }

    private const string executablePath = "C:\Windows\System32\cmd.exe";
    private const string workingDirectory = "C:\__TEMP\egov-validationclient";
    private const string url = "https://egovsigval-backend-a.bit.admin.ch";
    private const string username = "ABN-Kornacki";
    private const string password = "QA8aJNX6huXWsGB";
    private const string mandator = "FullQualified";
    private const int timeoutInSeconds = 20;

    public ValidateSignatureActivityCommandModule () : base ("ValidateSignature:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<ModuleParameters> (commandActivity);
      var documentExpression = parameters.Document;

      var document = documentExpression?.Invoke<Document> (commandActivity, commandActivity.WorkItem);
      if (!document.HasReadAccess())
      {
        throw new ActivityCommandException (
                $"The current user does not have permission to read document '{document.DisplayName}'.")
            .WithUserMessage (
                ModuleLocalization.ReadAccessDeniedError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, document.DisplayName));
      }

      if (document == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(ModuleParameters.Document)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.ParameterNullError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, nameof(ModuleParameters.Document)));
      }

      var activeContent = document.ActiveContent;
      if (activeContent == null)
      {
        throw new ActivityCommandException (
                $"Document '{document.DisplayName}' must have active content in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.MissingContentError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, document.DisplayName));
      }

      var tempDir = NamedFolder.GetTempFolder()?.FolderPath ?? Path.GetTempPath();
      var pathInbox = Path.Combine (tempDir, "inbox");
      var pathOutbox = Path.Combine (tempDir, "outbox");

      try
      {
        try
        {
          CreateDirectory(pathInbox);
          CreateDirectory(pathOutbox);
        }
        catch (Exception ex)
        {
          throw new ActivityCommandException($"The setup of the temporary directories failed in commandActivity '{commandActivity.ID}'.", ex).WithUserMessage(
              ModuleLocalization.FolderSetupError.ToLocalizedName().FormatWith(commandActivity.DisplayName));
        }

        var cleanDocumentName = CleanName (document.DisplayName);
        var documentPath = Path.Combine (pathInbox, cleanDocumentName);
        var reportPath = Path.Combine (pathOutbox, $"Validation_Report_{cleanDocumentName}");

        try
        {
          using (var fs = new FileStream (documentPath, FileMode.Create))
          {
            document.ActiveContent.GetContent().CreateStream().CopyTo (fs);
          }
        }
        catch (Exception ex)
        {
          throw new ActivityCommandException ($"Downloading the pdf failed in commandActivity '{commandActivity.ID}'.", ex).WithUserMessage (
              ModuleLocalization.DownloadPdfError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }

        var paramText = $"/c .\\validate.bat -u {url} -un {username} -pw {password} -m {mandator} -f {documentPath} -c -o {reportPath}";
        var timeout = TimeSpan.FromSeconds (timeoutInSeconds);
        var commandLineExecutorResult = CommandLineExecutor.RunCommand (executablePath, paramText, true, workingDirectory, timeout);
        var standardOutput = commandLineExecutorResult.StandardOutput;
        if (commandLineExecutorResult.ExitCode != 0)
        {
          throw new ActivityCommandException (
                  $"Error \"'{standardOutput}'\" occurred during the execution of the signature validation in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.SignatureValidationError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }

        FileHandlingContainer file;
        if (commandActivity.WorkItem.GetPublicDomainObjectType() == typeof (FileCase))
        {
          file = (FileHandlingContainer)((FileCase)commandActivity.WorkItem).ParentObject;
        }
        else
        {
          file = (FileHandlingContainer)commandActivity.WorkItem;
        }

        var validationReport = Document.NewObject (file);
        validationReport.PrimaryContent.Extension = "pdf";
        validationReport.PrimaryContent.Name = $"Resultat Signaturvalidierung {document.DisplayName.Replace (".pdf", "")}";
        validationReport.PrimaryContent.MimeType = MediaTypeNames.Application.Pdf;

        try
        {
          using (
              FileStream fs = new FileStream ($"{reportPath}", FileMode.Open))
          {
            validationReport.PrimaryContent.SetContent (fs.ReadAllBytes());
          }

          if (commandActivity.EffectiveWorkItemType.Name == "FileCase")
          {
            var fileCaseContent = FileCaseContent.NewObject();
            fileCaseContent.Document = validationReport;
            ((FileCase)commandActivity.WorkItem).FileCaseContents.Add (fileCaseContent);
          }
        }
        catch (Exception ex)
        {
          throw new ActivityCommandException ($"Uploading the validation report failed in commandActivity '{commandActivity.ID}'.", ex).WithUserMessage (
              ModuleLocalization.UploadReportError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }
      }
      finally
      {
        if (Exists (pathInbox))
          Delete (pathInbox, true);

        if (Exists (pathOutbox))
          Delete (pathOutbox, true);
      }

      return true;
    }

    private static string CleanName (string name)
    {
      var str = string.Join ("", name.Split (Path.GetInvalidFileNameChars()));
      return Regex.Replace (str, @"\s+", "_").Trim();
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<ModuleParameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);

        if (commandActivity.EffectiveWorkItemType == typeof (File)
            && commandActivity.EffectiveWorkItemType == typeof (Incoming)
            && commandActivity.EffectiveWorkItemType == typeof (FileCase))
        {
          throw new ActivityCommandException ($"The WorkItem must be of type File, Incoming or FileCase in activity '{commandActivity.DisplayName}'.")
              .WithUserMessage (ModuleLocalization.InvalidWorkItemTypeError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName));
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