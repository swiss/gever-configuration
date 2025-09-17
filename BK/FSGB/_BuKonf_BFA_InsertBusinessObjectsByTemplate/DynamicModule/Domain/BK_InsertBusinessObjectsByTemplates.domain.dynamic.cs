// Name: Insert Business Object Template
// Version: 2.0.0
// Datum: 09.10.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu; Ursprünglich LHIND - philipp.roessler@gs-ejpd.admin.ch

using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Remotion.Logging;
using Rubicon.Domain;
using Rubicon.Utilities;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using System;
using ActaNova.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Multilingual.Extensions;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities.Security;

namespace PoC.TestingOnly.InsertBusinessObjectsByTemplates
{
  [LocalizationEnum]
  public enum LocalizedUserMessages
  {
    [De ("100: Das Objekt dann nicht bearbeitet werden.")]
    NotEditable,

    [De ("200: Das Geschäftsobjekt der Aktivität \"{0}\" hat einen unzulässigen Typ.")]
    WrongBusinessObject,

    [De ("300: Auf Basis der ReferenceID konnte keine gültige Vorlage gefunden werden.")]
    NoTemplateFound,

    [De ("400: Es konnte kein gültiges Parent-Objekt gefunden werden.")]
    NoParentFound,

    [De ("500: Das Objekt kann nicht erstellt werden.")]
    TargetCannotCreateError
  }

  public class GeverBundInsertBusinessObjectByTemplate : ActivityCommandModule
  {
    private static readonly ILog s_logger = LogManager.GetLogger (typeof (GeverBundInsertBusinessObjectByTemplate));

    public GeverBundInsertBusinessObjectByTemplate () :
        base ("GEVERBund_InsertBusinessObjectByTemplate:ActivityCommandClassificationType")
    {
    }

    public override bool Execute (CommandActivity commandActivity)
    {
      FileHandlingObject hostObject;
      var onError = commandActivity.GetOrCreateFlowVariable ("onError_InsertBusinessObjectByTemplate");
      var createdObjectRefId = commandActivity.GetOrCreateFlowVariable ("createdObjectRefId_InsertBusinessObjectByTemplate");

      try
      {
        switch (commandActivity.WorkItem.ID.ClassID)
        {
          case "FileCase":
            hostObject = (FileCase)commandActivity.WorkItem;
            break;
          case "File":
            hostObject = (File)commandActivity.WorkItem;
            break;
          case "Incoming":
            hostObject = (Incoming)commandActivity.WorkItem;
            break;
          default:
            throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.WrongBusinessObject
                .ToLocalizedName()
                .FormatWith (commandActivity));
        }

        var templateIdParam = string.Empty;

        if (commandActivity.Parameter != null)
        {
          templateIdParam = commandActivity.Parameter;

          // Loglevel Debug, um den ISCeco Betriebsrichtlinien zu entsprechen.
          s_logger.Debug ($"Parameter: {templateIdParam}");
        }
        else
        {
          onError.SetValue ("100: Invalid Argument - Ungültiger Parameter");
        }

        var templateObject = new ReferenceHandle<BaseFileHandlingObjectTemplate> (templateIdParam).GetObject();

        if (!hostObject.CanEdit (true))
          throw new ActivityCommandException ($"Workitem '{hostObject.DisplayName}' is not editable.")
              .WithUserMessage (LocalizedUserMessages.NotEditable
                  .ToLocalizedName().FormatWith (commandActivity));

        switch (templateObject.ID.ClassID)
        {
          case "FileTemplate":
            var fileObject = File.NewObject (hostObject as File, null, templateObject as FileTemplate);

            createdObjectRefId.SetValue (fileObject.ToHasReferenceID().ReferenceID);

            break;
          case "FileCaseTemplate":
            if (!AccessControlUtility.HasConstructorAccess (typeof (FileCase)))
              throw new ActivityCommandException ($"User has no contructor access for '{typeof (FileCase).FullName}' in activity '{commandActivity.ID}'.")
                  .WithUserMessage (
                      LocalizedUserMessages.TargetCannotCreateError.ToLocalizedName().FormatWith (commandActivity));

            var fileCaseObject = FileCase.NewObject (hostObject as File, null, null, templateObject as FileCaseTemplate);
            var newTitle = $"{fileCaseObject.GetMultilingualValue (fc => fc.Title)} - {hostObject.DisplayName} - FileCase";

            fileCaseObject.SetMultilingualValue (fc => fc.Title, newTitle);
            createdObjectRefId.SetValue (fileCaseObject.ToHasReferenceID().ReferenceID);

            break;
          default:
            throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.WrongBusinessObject
                .ToLocalizedName()
                .FormatWith (commandActivity));
        }

        onError.SetValue ("0");
      }
      catch (Exception ex)
      {
        s_logger.Error ("Error Handling - Catch Part");
        onError.SetValue ("-1: Someting went wrong.");
        s_logger.Error (ex.Message, ex);
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      return commandActivity.Parameter != null
          ? null
          : "Bei der Aktivität muss ein Parameter mit der Referenz-ID der einzufügenden Geschäftsobjektvorlage übergeben werden.";
    }
  }

  public class GeverBundInsertBusinessObjectToParentByTemplate : ActivityCommandModule
  {
    private static readonly ILog s_logger = LogManager.GetLogger (typeof (GeverBundInsertBusinessObjectToParentByTemplate));

    public GeverBundInsertBusinessObjectToParentByTemplate () : base ("GEVERBund_InsertBusinessObjectToParentByTemplate:ActivityCommandClassificationType")
    {
    }

    public override bool Execute (CommandActivity commandActivity)
    {
      FileHandlingObject hostObject;
      FileHandlingObject parentObject;
      var onError = commandActivity.GetOrCreateFlowVariable ("onError_InsertBusinessObjectToParentByTemplate");
      var createdObjectRefId = commandActivity.GetOrCreateFlowVariable ("createdObjectRefId_InsertBusinessObjectToParentByTemplate");

      try
      {
        switch (commandActivity.WorkItem.ID.ClassID)
        {
          case "FileCase":
            hostObject = (FileCase)commandActivity.WorkItem;
            parentObject = ((FileCase)hostObject).BaseFile;
            break;
          case "File":
            hostObject = (File)commandActivity.WorkItem;
            if (((File)hostObject).ParentFile != null)
              parentObject = ((File)hostObject).ParentFile;
            else
              throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.NoParentFound
                  .ToLocalizedName()
                  .FormatWith (commandActivity));
            break;
          case "Incoming":
            hostObject = (Incoming)commandActivity.WorkItem;
            if (((Incoming)hostObject).BaseFile != null)
              parentObject = ((Incoming)hostObject).BaseFile;
            else
              throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.NoParentFound
                  .ToLocalizedName()
                  .FormatWith (commandActivity));
            break;
          default:
            throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.WrongBusinessObject
                .ToLocalizedName()
                .FormatWith (commandActivity));
        }

        var templateIdParam = string.Empty;

        if (commandActivity.Parameter != null)
        {
          templateIdParam = commandActivity.Parameter;

          // Loglevel Debug, um den ISCeco Betriebsrichtlinien zu entsprechen.
          s_logger.Debug ($"Parameter: {templateIdParam}");
        }
        else
        {
          onError.SetValue ("100: Invalid Argument - Ungültiger Parameter");
        }

        var templateObject = new ReferenceHandle<BaseFileHandlingObjectTemplate> (templateIdParam).GetObject();

        if (!hostObject.CanEdit (true))
          throw new ActivityCommandException ($"Workitem '{hostObject.DisplayName}' is not editable.")
              .WithUserMessage (LocalizedUserMessages.NotEditable
                  .ToLocalizedName().FormatWith (commandActivity));

        switch (templateObject.ID.ClassID)
        {
          case "FileTemplate":
            var fileObject = File.NewObject (parentObject as File, null, templateObject as FileTemplate);
            createdObjectRefId.SetValue (fileObject.ToHasReferenceID().ReferenceID);

            break;
          case "FileCaseTemplate":
            if (!AccessControlUtility.HasConstructorAccess (typeof (FileCase)))
              throw new ActivityCommandException ($"User has no contructor access for '{typeof (FileCase).FullName}' in activity '{commandActivity.ID}'.")
                  .WithUserMessage (
                      LocalizedUserMessages.TargetCannotCreateError.ToLocalizedName().FormatWith (commandActivity));

            var fileCaseObject = FileCase.NewObject (parentObject as File, null, null, templateObject as FileCaseTemplate);
            var newTitle = $"{fileCaseObject.GetMultilingualValue (fc => fc.Title)} - {parentObject.DisplayName} - FileCase";

            fileCaseObject.SetMultilingualValue (fc => fc.Title, newTitle);
            createdObjectRefId.SetValue (fileCaseObject.ToHasReferenceID().ReferenceID);

            break;
          default:
            throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.WrongBusinessObject
                .ToLocalizedName()
                .FormatWith (commandActivity));
        }

        onError.SetValue ("0");
      }
      catch (Exception ex)
      {
        s_logger.Error ("Error Handling - Catch Part");
        onError.SetValue ("-1: Someting went wrong");
        s_logger.Error (ex.Message, ex);
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      return commandActivity.Parameter != null
          ? null
          : "Bei der Aktivität muss ein Parameter mit der Referenz-ID der einzufügenden Geschäftsobjektvorlage übergeben werden.";
    }
  }

  public class GeverBundInsertDocumentByTemplate : ActivityCommandModule
  {
    private static readonly ILog s_logger = LogManager.GetLogger (typeof (GeverBundInsertDocumentByTemplate));

    public GeverBundInsertDocumentByTemplate () : base ("GEVERBund_InsertDocumentByTemplate:ActivityCommandClassificationType")
    {
    }

    public override bool Execute (CommandActivity commandActivity)
    {
      var onError = commandActivity.GetOrCreateFlowVariable ("onError_InsertDocumentByTemplate");

      try
      {
        var templateIdParam = string.Empty;

        if (commandActivity.Parameter != null)
        {
          templateIdParam = commandActivity.Parameter;

          // Loglevel Debug, um den ISCeco Betriebsrichtlinien zu entsprechen.
          s_logger.Debug ("Parameter: " + templateIdParam);
        }
        else
        {
          onError.SetValue ("100: Invalid Argument - Ungültiger Parameter");
        }

        var templateObject = new ReferenceHandle<DocumentTemplate> (templateIdParam).GetObject();

        if (templateObject == null)
          throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.NoTemplateFound
              .ToLocalizedName()
              .FormatWith (templateObject));

        if (!AccessControlUtility.HasConstructorAccess (typeof (Document)))
          throw new ActivityCommandException ($"User has no contructor access for '{typeof (Document).FullName}' in activity '{commandActivity.ID}'.")
              .WithUserMessage (
                  LocalizedUserMessages.TargetCannotCreateError.ToLocalizedName().FormatWith (commandActivity));

        switch (commandActivity.WorkItem.ID.ClassID)
        {
          case "FileCase":
            var fileCaseObject = (FileCase)commandActivity.WorkItem;

            if (!fileCaseObject.CanEdit (true))
              throw new ActivityCommandException ("FileCase '" + fileCaseObject.DisplayName + "' is not editable.")
                  .WithUserMessage (LocalizedUserMessages.NotEditable
                      .ToLocalizedName().FormatWith (commandActivity));

            var newDocument = Document.NewObject (fileCaseObject.GetParentFileHandlingContainer(), templateObject);
            fileCaseObject.AddFileCaseContent (newDocument);

            break;
          case "File":
            var fileObject = (File)commandActivity.WorkItem;

            if (!fileObject.CanEdit (true))
              throw new ActivityCommandException ("File '" + fileObject.DisplayName + "' is not editable.")
                  .WithUserMessage (LocalizedUserMessages.NotEditable
                      .ToLocalizedName().FormatWith (commandActivity));

            Document.NewObject (fileObject, templateObject);

            break;
          case "Incoming":
            var incomingObject = (Incoming)commandActivity.WorkItem;

            if (!incomingObject.CanEdit (true))
              throw new ActivityCommandException ("Incoming '" + incomingObject.DisplayName + "' is not editable.")
                  .WithUserMessage (LocalizedUserMessages.NotEditable
                      .ToLocalizedName().FormatWith (commandActivity));

            Document.NewObject (incomingObject, templateObject);

            break;
          default:
            throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.WrongBusinessObject
                .ToLocalizedName()
                .FormatWith (commandActivity));
        }

        onError.SetValue ("0");
      }
      catch (Exception ex)
      {
        s_logger.Error ("Error Handling - Catch Part");
        onError.SetValue ("-1: Someting went wrong");
        s_logger.Error (ex.Message, ex);
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      return commandActivity.Parameter != null
          ? null
          : "Bei der Aktivität muss ein Parameter mit der Referenz-ID der anzuwendenen Dokumentvorlage übergeben werden.";
    }
  }
}