// Name: DeepL-Integration
// Version: 1.0.0
// Datum: 18.10.2024
// Autor: RUBICON IT

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using ActaNova.Domain;
using ActaNova.Web.Classes;
using ActaNova.Web.Controls;
using ActaNova.Web.Controls.Hierarchy;
using Newtonsoft.Json;
using Remotion.Data.DomainObjects;
using Remotion.Globalization;
using Remotion.Web.UI.Controls;
using Rubicon.Dms;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.ObjectBinding.Forms.Web.Controls;
using Rubicon.Utilities;
using Rubicon.Utilities.Autofac;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Mime;
using Rubicon.Web.Utilities.Autofac;

namespace Rubicon.Gever.Bund.Configs.Configs.Sample
{
  public class TranslatorFormCommandModule : BocListFormCommandModule
  {
    protected override void Load (CachingFactory<ControlPropertyBindingKey, IList<BocListFormCommand>, GetFormCommands>.Registrator registrator)
    {
      Register(registrator).For<FileContentHierarchyBocList>()
          // ReSharper disable once SuspiciousTypeConversion.Global
          .WhenBoundTo<FileHandlingContainer>(o => o.FileContentHierarchy)
          .AddCommand<DeepLTranslatorDE>(RequiredSelection.ExactlyOne)
          .AddCommand<DeepLTranslatorFR>(RequiredSelection.ExactlyOne)
          .AddCommand<DeepLTranslatorIT>(RequiredSelection.ExactlyOne)
          .AddCommand<DeepLTranslatorEN>(RequiredSelection.ExactlyOne);

      registrator.Builder.Register<DeepLHttpClient>().SingletonScoped();
    }
  }

  [JsonObject]
  public class DeepLDocument
  {
    [JsonProperty("document_id")]
    public string ID { get; set; }
    [JsonProperty("document_key")]
    public string Key { get; set; }

    public bool ShouldSerializeID () => false;
  }

  [JsonObject]
  public class DeepLDocumentTranslationStatusResult
  {
    [JsonProperty ("document_id")]
    public string ID;

    [JsonProperty ("status")]
    public string Status;

    [JsonProperty ("seconds_remaining")]
    public int SecondsRemaining;

    [JsonProperty ("billed_characters")]
    public int BilledCharacters;

    [JsonProperty ("error_message")]
    public string ErrorMessage;
  }

  public class DeepLHttpClient : HttpClient
  {
    private readonly Uri _deepLServiceBaseAddress = new Uri("https://api.deepl.com/");
    private const string _deepLAuthKey = "DeepL-Auth-Key bb4d9298-9aa1-4f2a-9bdc-6ae0c8b0be74";

    public DeepLHttpClient () : base (HttpClientHandlerWithProxy())
    {
      BaseAddress = _deepLServiceBaseAddress;
      DefaultRequestHeaders.Add("Authorization", _deepLAuthKey);
    }

    private static HttpClientHandler HttpClientHandlerWithProxy ()
    {
      return new HttpClientHandler
      {
          Proxy = new WebProxy("{{PROXY_URL}}"),
          UseProxy = true
      };
    }
  }

  public static class MimeTypeToFileExtensionMapping
  {
    public static string GetFileExtension (string mimeType)
    {
      var mimeMapping = new Dictionary<string, string>();
      mimeMapping.Add("application/pdf", ".pdf");
      mimeMapping.Add("text/plain", ".txt");
      mimeMapping.Add(MimeTypes.ApplicationMSWord, ".doc");
      mimeMapping.Add(MimeTypes.ApplicationOpenOfficeDocument, ".docx");
      mimeMapping.Add("application/vnd.ms-excel", ".xls");
      mimeMapping.Add("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx");
      return mimeMapping[mimeType];
    }
  }

  [LocalizationEnum]
  public enum DeepLLocalization
  {
    [De("Dokument übersetzen (DE)")]
    CommandTitleDE,
    [De("Dokument übersetzen (FR)")]
    CommandTitleFR,
    [De("Dokument übersetzen (IT)")]
    CommandTitleIT,
    [De("Dokument übersetzen (EN)")]
    CommandTitleEN,
    [De("Beim Hochladen des Dokuments ist ein Fehler aufgetreten. (Antwort Code: {0})")]
    ErrorAtDocumentUpload,
    [De("Beim Hochladen des Dokuments ist ein Fehler aufgetreten.")]
    ErrorAtDocumentUploadGeneric,
    [De("Die Übersetzung wurde nicht rechtzeitig fertig.")]
    TranslationTimeout,
    [De("Bei der Übersetzung ist ein Fehler aufgetreten: {0}")]
    ErrorAtTranslation,
    [De("Beim Herunterladen des übersetzten Dokument ist ein Fehler aufgetreten. (Antwort Code: {0})")]
    ErrorDownloadingTranslatedDocument,
    [De("Beim Herunterladen des übersetzten Dokument ist ein Fehler aufgetreten.")]
    ErrorDownloadingTranslatedDocumentGeneric,
    [De("Automatisch übersetzt mit DeepL")]
    DocumentRemarkTranslated
  }

  public class DeepLTranslatorDE : DeepLTranslatorBase
  {
    protected override string TargetLanguage => "DE";

    public DeepLTranslatorDE ()
    {
      Title = DeepLLocalization.CommandTitleDE.ToLocalizedName();
    }
  }
  public class DeepLTranslatorFR : DeepLTranslatorBase
  {
    protected override string TargetLanguage => "FR";

    public DeepLTranslatorFR ()
    {
      Title = DeepLLocalization.CommandTitleFR.ToLocalizedName();
    }
  }
  public class DeepLTranslatorIT : DeepLTranslatorBase
  {
    protected override string TargetLanguage => "IT";

    public DeepLTranslatorIT ()
    {
      Title = DeepLLocalization.CommandTitleIT.ToLocalizedName();
    }
  }
  public class DeepLTranslatorEN : DeepLTranslatorBase
  {
    protected override string TargetLanguage => "EN";

    public DeepLTranslatorEN ()
    {
      Title = DeepLLocalization.CommandTitleEN.ToLocalizedName();
    }
  }

  public abstract class DeepLTranslatorBase : ActaNovaFormCommand
  {
    private DeepLHttpClient HttpClient => Containers.Dynamic.Resolve<DeepLHttpClient>();

    protected abstract string TargetLanguage { get; }

    protected DeepLTranslatorBase ()
    {
      Type = CommandType.Event;
      // ReSharper disable DoNotCallOverridableMethodsInConstructor
      EventCommand.RequiresSynchronousPostBack = false;
      SaveSignature = false;
      SupportsMultiSelection = false;
      // ReSharper restore DoNotCallOverridableMethodsInConstructor
      RequiresEditableObject = true;
    }

    protected override FormCommandStateType ApplyAll ()
    {
      var deepLDocument = UploadDocument();
      
      var pollingStart = DateTime.Now;
      DeepLDocumentTranslationStatusResult status;
      while (true)
      {
        Thread.Sleep(2000);
        status = CheckDocumentStatus(deepLDocument);

        if (string.Compare(status.Status, "done", StringComparison.CurrentCultureIgnoreCase) == 0)
          break;

        if (string.Compare (status.Status, "error", StringComparison.CurrentCultureIgnoreCase) == 0)
          throw new ApplicationException ("Translation error: " + status.ErrorMessage)
              .WithUserMessage(string.Format(DeepLLocalization.ErrorAtTranslation.ToLocalizedName(), status.ErrorMessage));

        if (pollingStart.AddSeconds (30) < DateTime.Now)
          throw new ApplicationException ("Translation didn't finish in time.")
              .WithUserMessage(DeepLLocalization.TranslationTimeout.ToLocalizedName());
      }

      DownloadTranslatedDocument(deepLDocument);

      return FormCommandStateType.Success;
    }

    private DeepLDocument UploadDocument ()
    {
      try
      {
        var form = new MultipartFormDataContent();
        form.Add (new StringContent (TargetLanguage), "target_lang");

        var document = ((Document)BusinessObject);
        var fileName = document.ActiveContent.GetFullName();
        HttpResponseMessage response;
        using (var handle = document.ActiveContent.GetContent())
        using (var stream = handle.CreateStream())
        {
          var fileStreamContent = new StreamContent (stream);
          fileStreamContent.Headers.ContentType = new MediaTypeHeaderValue ("application/octet-stream");
          form.Add (fileStreamContent, "file", "@" + fileName);
          response = HttpClient.PostAsync ("/v2/document", form).Result;
        }

        if (!response.IsSuccessStatusCode)
          throw new ApplicationException ($"Error response received ({response.StatusCode}) by uploading document to DeepL.")
              .WithUserMessage (string.Format (DeepLLocalization.ErrorAtDocumentUpload.ToLocalizedName(), response.StatusCode));

        var responseBody = response.Content.ReadAsStringAsync().Result;
        return JsonConvert.DeserializeObject<DeepLDocument> (responseBody);
      }
      catch (ApplicationException e)
      {
        throw;
      }
      catch (Exception e)
      {
        throw new ApplicationException($"Upload to DeepL failed with exception: {e.Message}")
            .WithUserMessage(string.Format(DeepLLocalization.ErrorAtDocumentUploadGeneric.ToLocalizedName(), e.Message));
      }
    }

    private DeepLDocumentTranslationStatusResult CheckDocumentStatus (DeepLDocument deepLDocument)
    {
      var deepLDocumentJson = JsonConvert.SerializeObject (deepLDocument);
      var requestContent = new StringContent(deepLDocumentJson);
      requestContent.Headers.ContentType = new MediaTypeHeaderValue ("application/json");
      var documentStatusCheckResponse = HttpClient.PostAsync ("/v2/document/" + deepLDocument.ID, requestContent).Result;
      var responseBody = documentStatusCheckResponse.Content.ReadAsStringAsync().Result;
      var statusCheckResult = JsonConvert.DeserializeObject<DeepLDocumentTranslationStatusResult> (responseBody);
      return statusCheckResult;
    }

    private void DownloadTranslatedDocument (DeepLDocument deepLDocument)
    {
      try
      {
        var deepLDocumentJson = JsonConvert.SerializeObject (deepLDocument);
        var requestContent = new StringContent (deepLDocumentJson);
        requestContent.Headers.ContentType = new MediaTypeHeaderValue ("application/json");
        var documentDownloadResponse = HttpClient.PostAsync ("/v2/document/" + deepLDocument.ID + "/result", requestContent).Result;
        if (!documentDownloadResponse.IsSuccessStatusCode)
          throw new ApplicationException ("Error at downloading the translated document.")
              .WithUserMessage (string.Format (DeepLLocalization.ErrorDownloadingTranslatedDocument.ToLocalizedName(), documentDownloadResponse.StatusCode));

        var translatedDocument = Document.NewObject ((FileHandlingContainer)DataSourceBusinessObject);
        translatedDocument.Remark = DeepLLocalization.DocumentRemarkTranslated.ToLocalizedName();
        translatedDocument.ActiveContent.MimeType = documentDownloadResponse.Content.Headers.ContentType.MediaType;
        translatedDocument.ActiveContent.Extension =
            MimeTypeToFileExtensionMapping.GetFileExtension (documentDownloadResponse.Content.Headers.ContentType.MediaType);
        translatedDocument.ActiveContent.Name = ((Document)BusinessObject).ActiveContent.Name;
        translatedDocument.ActiveContent.SetContent (documentDownloadResponse.Content.ReadAsByteArrayAsync().Result);

        ClientTransaction.Current.Commit();
      }
      catch (ApplicationException e)
      {
        throw;
      }
      catch (Exception e)
      {
        throw new ApplicationException($"Downloading the translated Document failed with exception: {e.Message}")
            .WithUserMessage(string.Format(DeepLLocalization.ErrorDownloadingTranslatedDocumentGeneric.ToLocalizedName(), e.Message));
      }
    }
  }
}
