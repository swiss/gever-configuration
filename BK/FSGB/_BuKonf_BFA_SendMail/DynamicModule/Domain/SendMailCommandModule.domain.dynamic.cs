// Name: Mail Versand (SendMail)
// Version: 2.0.0
// Datum: 09.05.2025
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu - Ursprünglich: Dimitrij Zaks - dimitrij.zaks@lhind.dlh.de - Lufthansa Industry Solutions GmbH

using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubicon.Utilities.Email;
using Rubicon.Workflow.Domain;
using System.Reflection;
using System.Text;
using ActaNova.Domain;
using Rubicon.Dms;
using System.Net.Mail;
using Rubicon.Utilities.Autofac;
using ActaNova.Domain.Classes.Configuration;
using Remotion.Data.DomainObjects;
using Remotion.Data.DomainObjects.ObjectBinding;
using Rubicon.Gever.Bund.DocumentEncryption.Core;
using Rubicon.Utilities.IO;
using File = ActaNova.Domain.File;

namespace Bukonf.Gever.Domain
{
  [VersionInfo ("2.0.0", "09.05.2025", "RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu - Ursprünglich Dimitrij Zaks LHIND")]
  public class SendMailCommandModule : GeverActivityCommandModule
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

    [LocalizationEnum]
    private enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Für Parameter \"{1}\" wurde kein Wert gefunden.")]
      ExpressionReturnedNoValue,

      [De ("Befehlsaktivität \"{0}\": Beim Anhängen der Datei \"{1}\" ist ein Fehler aufgetreten.")]
      AttachmentError,

      [De ("Befehlsaktivität \"{0}\": Fehlermeldung : {1}.")]
      GeneralError,

      [De ("Befehlsaktivität \"{0}\": Beim Verarbeiten der E-Mail Parameter ist ein Fehler aufgetreten.")]
      ParameterProcessingError,

      [De ("Befehlsaktivität \"{0}\": Beim Versand der E-Mail ohne Ablage ist ein Fehler aufgetreten.")]
      ErrorSendingMailWithoutRecording,

      [De ("Befehlsaktivität \"{0}\": Beim Versand der E-Mail mit Ablage ist ein Fehler aufgetreten.")]
      ErrorSendingMailWithRecording,

      [De ("Befehlsaktivität \"{0}\": Die E-Mail-Adresse \"{1}\" ist ungültig.")]
      InvalidEmailAddress,

      [De ("Befehlsaktivität \"{0}\": Der Ausdruck für die Anhänge konnte nicht verarbeitet werden.")]
      AttachmentExpressionNullError,

      [De ("Befehlsaktivität \"{0}\": Beim Verarbeiten der Anhänge ist ein Fehler aufgetreten.")]
      AttachmentProcessingError,

      [De ("Befehlsaktivität \"{0}\": Beim Verarbeiten der Links ist ein Fehler aufgetreten.")]
      ProcessingLinksError,

      [De ("Befehlsaktivität \"{0}\": Version: \"{1}\" - Datum: \"{2}\" - Autor \"{3}\"")]
      ShowVersion,
    }

    public class Parameters : GeverActivityCommandParameters
    {
      // Do not remove unused parameters or change type of used parameters in version 2.0.0 to preserve backward compatibility with LHIND Version
      [Required]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Sender { get; set; }

      [Required]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Recipient { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression ReplyTo { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression RecipientCC { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression RecipientBCC { get; set; }

      [Required]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Subject { get; set; }

      [Required]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Body { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Priority { get; set; }

      [ExpectedExpressionType (typeof (Object))]
      public TypedParsedExpression Attachements { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression ExecuteAfterCommit { get; set; }

      [ExpectedExpressionType (typeof (bool?))]
      public TypedParsedExpression Deactivate { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression DefaultRecipient { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression SmptProvider { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression RecordSentMail { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression ShowVersion { get; set; }
    }

    public SendMailCommandModule () : base ("SendMail:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<Parameters> (commandActivity);
      var executeAfterCommit = parameters.ExecuteAfterCommit != null && parameters.ExecuteAfterCommit.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      if (executeAfterCommit)
      {
        commandActivity.Committed += (sender, e) => SendMail (commandActivity, parameters);
      }
      else
      {
        SendMail (commandActivity, parameters);
      }

      return true;
    }

    private void SendMail (CommandActivity commandActivity, Parameters parameters)
    {
      try
      {
        var recordSentMail = parameters.RecordSentMail != null && parameters.RecordSentMail.Invoke<bool> (commandActivity, commandActivity.WorkItem);
        var emailParameters = GetEmailParameters (commandActivity, parameters);

        int nonAdminRecipientCount = emailParameters.Recipients
            .Count (r => !r.Address.EndsWith ("admin.ch", StringComparison.OrdinalIgnoreCase));

        if (recordSentMail || emailParameters.Attachments.Any() || nonAdminRecipientCount > 0)
        {
          SendMailWithRecording (commandActivity, emailParameters);
        }
        else
        {
          SendMailWithoutRecording (commandActivity, emailParameters);
        }
      }
      catch (ActivityCommandException)
      {
        throw;
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException (
                $"CommandActivity \"{commandActivity.DisplayName}\": Error Message: {ex.Message}.")
            .WithUserMessage (
                ModuleLocalization.GeneralError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, ex.Message));
      }
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<Parameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);
        ShowVersion (commandActivity, parameters);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    private void ShowVersion (CommandActivity commandActivity, Parameters parameters)
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
                  .ShowVersion.ToLocalizedName().FormatWith (commandActivity.DisplayName, attr.Version, attr.Date, attr.Author));
        }
      }
    }

    private EmailParameters GetEmailParameters (CommandActivity commandActivity, Parameters parameters)
    {
      try
      {
        var emailParameters = new EmailParameters();
        emailParameters.From =
            new EmailProviderAddress (
                GetValidatedEmailAddress (ProcessStringExpression (parameters.Sender, commandActivity, nameof(parameters.Sender)), commandActivity).Address);

        // Do not use ProcessStringExpression for ReplyTo because it is not required
        if (emailParameters.ReplyTo != null)
        {
          emailParameters.ReplyTo =
              new EmailProviderAddress (
                  GetValidatedEmailAddress (parameters.ReplyTo.Invoke<string> (commandActivity, commandActivity.WorkItem), commandActivity).Address);
        }

        emailParameters.Recipients =
            ComputeRecipients (ProcessStringExpression (parameters.Recipient, commandActivity, nameof(parameters.Recipient)), commandActivity);
        var linkPlaceHolderProcessor = new LinkPlaceHolderProcessor();
        var unprocessedSubject = ProcessStringExpression (parameters.Subject, commandActivity, nameof(parameters.Subject));
        emailParameters.Subject = linkPlaceHolderProcessor.ProcessLinks (unprocessedSubject, commandActivity);
        var unprocessedBody = ProcessStringExpression (parameters.Body, commandActivity, nameof(parameters.Body));
        emailParameters.Body = linkPlaceHolderProcessor.ProcessLinks (unprocessedBody, commandActivity);
        emailParameters.Attachments = new AttachmentProcessor().ProcessAttachments (parameters, commandActivity);

        return emailParameters;
      }
      catch (ActivityCommandException)
      {
        throw;
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException (
                $"CommandActivity \"{commandActivity.DisplayName}\": An error occured while processing the E-Mail parameters. {ex.Message}")
            .WithUserMessage (ModuleLocalization.ParameterProcessingError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }
    }

    private void SendMailWithoutRecording (CommandActivity commandActivity, EmailParameters emailParameters)
    {
      try
      {
        var emailProvider =
            DynamicCachingFactoryModule<EmailProviderModuleKey, Func<IEmailProvider>, GetEmailProvider>.Value (nameof(SendMailCommandModule),
                EmailProviderCapability.HtmlBody);

        emailProvider.Send (from: emailParameters.From,
            userEmail: emailParameters.ReplyTo,
            replyTo: null,
            toRecipients: emailParameters.Recipients,
            bccRecipients: null,
            ccRecipients: null,
            subject: emailParameters.Subject,
            body: new HtmlEmailBody (emailParameters.Body, Encoding.UTF8),
            priority: EmailProviderMailPriority.Normal,
            attachments: emailParameters.Attachments,
            headers: null);
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException (
                $"CommandActivity \"{commandActivity.DisplayName}\": Error sending email without recording. {ex.Message}")
            .WithUserMessage (ModuleLocalization.ErrorSendingMailWithoutRecording.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }
    }

    private void SendMailWithRecording (CommandActivity commandActivity, EmailParameters emailParameters)
    {
      try
      {
        var outgoingEmailDocument = OutgoingEmailDocument.NewObject ((FileHandlingContainer)commandActivity.WorkItem);
        foreach (var recipient in emailParameters.Recipients)
        {
          var emailRecipient = EmailRecipient.NewObject (null, recipient.Address);
          outgoingEmailDocument.EmailRecipients.Add (emailRecipient);
        }

        outgoingEmailDocument.EmailFrom = emailParameters.From.Address;
        outgoingEmailDocument.EmailSubject = emailParameters.Subject;
        outgoingEmailDocument.EmailBody = emailParameters.Body;
        AddOutgoingEmailDocumentAttachments (emailParameters.Attachments, outgoingEmailDocument, commandActivity);
      }
      catch (ActivityCommandException)
      {
        throw;
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException (
                $"CommandActivity \"{commandActivity.DisplayName}\": Error sending email with recording. {ex.Message}")
            .WithUserMessage (ModuleLocalization.ErrorSendingMailWithRecording.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }
    }

    private string ProcessStringExpression (TypedParsedExpression expression, CommandActivity commandActivity, string paramName)
    {
      var expressionResult = expression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      if (expressionResult.IsNullOrEmptyOrBlank())
        throw new ActivityCommandException ($"CommandActivity {commandActivity.DisplayName}: The value for parameter {paramName} could not be parsed.")
            .WithUserMessage (ModuleLocalization.ExpressionReturnedNoValue.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, paramName));
      return expressionResult;
    }

    private IEnumerable<IEmailProviderAddress> ComputeRecipients (string recipients, CommandActivity commandActivity)
    {
      List<IEmailProviderAddress> recipientAddresses = new List<IEmailProviderAddress>();
      var recipientList = recipients.Split (';');
      foreach (var recipient in recipientList)
      {
        recipientAddresses.Add (new EmailProviderAddress (GetValidatedEmailAddress (recipient, commandActivity).Address));
      }

      return recipientList.Length > 0 ? recipientAddresses : Enumerable.Empty<IEmailProviderAddress>();
    }

    private MailAddress GetValidatedEmailAddress (string emailAddress, CommandActivity commandActivity)
    {
      try
      {
        return new MailAddress (emailAddress);
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException (
                $"CommandActivity \"{commandActivity.DisplayName}\": Invalid recipient address '{emailAddress}'. {ex.Message}")
            .WithUserMessage (
                ModuleLocalization.InvalidEmailAddress.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, emailAddress));
      }
    }

    private void AddOutgoingEmailDocumentAttachments (
        List<IEmailProviderAttachment> attachments,
        OutgoingEmailDocument outgoingEmailDocument,
        CommandActivity commandActivity)
    {
      foreach (var attachment in attachments)
      {
        try
        {
          var outgoingEmailAttachmentDocument = OutgoingEmailAttachmentDocument.NewObject (outgoingEmailDocument);

          outgoingEmailAttachmentDocument.PrimaryContent.MimeType = attachment.MimeType;
          outgoingEmailAttachmentDocument.PrimaryContent.Name = attachment.Name;
          outgoingEmailAttachmentDocument.PrimaryContent.Extension = attachment.Extension;

          using (var handle = new SimpleStreamAccessHandle (s => new MemoryStream (attachment.Content)))
          {
            outgoingEmailAttachmentDocument.PrimaryContent.SetContent (handle, attachment.Extension, attachment.MimeType);
          }
        }
        catch (Exception ex)
        {
          throw new ActivityCommandException (
                  $"CommandActivity \"{commandActivity.DisplayName}\": An error occured while adding the attachment {attachment.Name}.{ex.Message}")
              .WithUserMessage (ModuleLocalization.AttachmentError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, attachment.Name));
        }
      }
    }

    // Implementation from LHIND
    private class AttachmentProcessor
    {
      public List<IEmailProviderAttachment> ProcessAttachments (Parameters parameters, CommandActivity commandActivity)
      {
        try
        {
          var result = new List<IEmailProviderAttachment>();
          if (parameters.Attachements == null)
            return result;

          Object expressionResult = parameters.Attachements.Invoke (commandActivity, commandActivity.WorkItem);
          if (expressionResult == null)
          {
            throw new ActivityCommandException (
                    $"CommandActivity \"{commandActivity.DisplayName}\": The expression for the attachments returned null.")
                .WithUserMessage (
                    ModuleLocalization.AttachmentExpressionNullError.ToLocalizedName()
                        .FormatWith (commandActivity.DisplayName));
          }

          if (expressionResult.GetType() == typeof (List<Document>))
          {
            ((List<Document>)expressionResult).Where (d => !d.AsEncryptableDocument().IsContentEncrypted).ForEach (d => result.Add (ProcessAttachment (d)));
          }
          else if (expressionResult.GetType() == typeof (ObjectList<Document>))
          {
            ((ObjectList<Document>)expressionResult).Where (d => !d.AsEncryptableDocument().IsContentEncrypted)
                .ForEach (d => result.Add (ProcessAttachment (d)));
          }
          else if (expressionResult is Document document)
          {
            result.Add (ProcessAttachment (document));
          }
          else if (expressionResult.GetType() == typeof (List<FileContentObject>))
          {
            ((List<FileContentObject>)expressionResult).Select (o => (Document)o).Where (d => !d.AsEncryptableDocument().IsContentEncrypted)
                .ForEach (d => result.Add (ProcessAttachment (d)));
          }
          else if (expressionResult.GetType().ToString().StartsWith ("System.Linq.Enumerable+WhereSelectEnumerableIterator"))
          {
            ((IEnumerable<FileContentObject>)expressionResult).Select (o => (Document)o).Where (d => !d.AsEncryptableDocument().IsContentEncrypted)
                .ForEach (d => result.Add (ProcessAttachment (d)));
          }

          return result;
        }
        catch (ActivityCommandException)
        {
          throw;
        }
        catch (Exception ex)
        {
          throw new ActivityCommandException (
                  $"CommandActivity \"{commandActivity.DisplayName}\": Error processing attachments. {ex.Message}")
              .WithUserMessage (
                  ModuleLocalization.AttachmentProcessingError.ToLocalizedName()
                      .FormatWith (commandActivity.DisplayName));
        }
      }

      private IEmailProviderAttachment ProcessAttachment (Document doc)
      {
        var contentMimeType = doc.ActiveContent.MimeType;
        var Name = doc.ActiveContent.Name;
        var Extension = doc.ActiveContent.Extension;

        using (MemoryStream ms = new MemoryStream())
        {
          using (IStreamAccessHandle content = doc.ActiveContent.GetContent())
          {
            content.CreateStream().CopyTo (ms);
          }

          return new EmailProviderAttachment (ms.ToArray(), Name, contentMimeType, Extension);
        }
      }
    }

    // Implementation from LHIND
    private class LinkPlaceHolderProcessor
    {
      private const string Placeholder_parent = "[[Übergeordnetes Geschäftsobjekt]]";
      private const string Placeholder_self = "[[Geschäftsobjekt]]";
      private const string Placeholder_docs = "[[Dokumente]]";
      private const string Placeholder_parent_docs = "[[Übergeordnete Dokumente]]";

      public LinkPlaceHolderProcessor ()
      {
      }

      public string ProcessLinks (string text, CommandActivity commandActivity)
      {
        try
        {
          var result = text.Contains (Placeholder_self) ? text.Replace (Placeholder_self, ToLink (commandActivity.WorkItem, false)) : text;
          result = result.Contains (Placeholder_parent) ? result.Replace (Placeholder_parent, ToLink (GetParent (commandActivity.WorkItem), false)) : result;
          result = result.Contains (Placeholder_parent_docs)
              ? result.Replace (Placeholder_parent_docs, CreateDocLinks (GetParent (commandActivity.WorkItem)))
              : result;
          return result.Contains (Placeholder_docs) ? result.Replace (Placeholder_docs, CreateDocLinks (commandActivity.WorkItem)) : result;
        }
        catch (Exception ex)
        {
          throw new ActivityCommandException ($"CommandActivity \"{commandActivity.DisplayName}\": Error processing link placeholders. {ex.Message}")
              .WithUserMessage (ModuleLocalization.ProcessingLinksError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }
      }

      private string CreateDocLinks (BindableDomainObject workItem)
      {
        var link = "";
        var ob = GetDocumentsHost (workItem);
        if (ob == null)
        {
          return link;
        }

        if (ob.GetPublicDomainObjectType() == typeof (FileCase))
        {
          link = ToLinks (((FileCase)ob).FileCaseContents.Select (e => (BindableDomainObject)e), true);
        }
        else if (ob.GetPublicDomainObjectType() == typeof (File))
        {
          link = ToLinks (((File)ob).Documents.Select (e => (BindableDomainObject)e), true);
        }

        return link == "" ? "" : "<p>" + link + "</p>";
      }

      private static BindableDomainObject GetDocumentsHost (BindableDomainObject workItem)
      {
        if (workItem.GetPublicDomainObjectType() == typeof (FileCase))
          return workItem;

        if (workItem.GetPublicDomainObjectType() == typeof (File))
          return workItem;

        if (workItem.GetPublicDomainObjectType() == typeof (Incoming))
          return (((Incoming)workItem).ParentObject is BindableDomainObject) ? (BindableDomainObject)((Incoming)workItem).ParentObject : null;

        return null;
      }

      private string ToLinks (IEnumerable<BindableDomainObject> list, bool withLineBreak)
      {
        return list.Aggregate ("", (current, o) => current + ToLink (o, withLineBreak));
      }

      private static string ToLink (BindableDomainObject item, bool withLineBreak)
      {
        var link = "<a href=\"" + UrlProvider.Current.GetOpenWorkListItemUrl (item) + "\">" +
                   item.DisplayName + "</a>" + (withLineBreak ? "<br/>" : "");
        return link;
      }


      private BindableDomainObject GetParent (FileCase fc)
      {
        return fc.BaseFile;
      }

      private BindableDomainObject GetParent (File f)
      {
        return (BindableDomainObject)f.ParentObject;
      }

      private BindableDomainObject GetParent (Incoming inc)
      {
        return (BindableDomainObject)inc.ParentObject;
      }

      private BindableDomainObject GetParent (BindableDomainObject workItem)
      {
        if (workItem.GetPublicDomainObjectType() == typeof (File))
          return GetParent ((File)workItem);

        if (workItem.GetPublicDomainObjectType() == typeof (FileCase))
          return GetParent ((FileCase)workItem);

        if (workItem.GetPublicDomainObjectType() == typeof (Incoming))
          return GetParent ((Incoming)workItem);

        return null;
      }
    }

    private class EmailParameters
    {
      public EmailProviderAddress From { get; set; }
      public EmailProviderAddress ReplyTo { get; set; }
      public IEnumerable<IEmailProviderAddress> Recipients { get; set; }
      public string Subject { get; set; }
      public string Body { get; set; }
      public List<IEmailProviderAttachment> Attachments { get; set; }
    }
  }
}