using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Bukonf.Gever.Domain
{
    public class DeeplClient : IDisposable
    {
        private const HttpStatusCode HttpStatusCodeTooManyRequests = (HttpStatusCode)429;

        /// <summary>HTTP status code returned by DeepL API to indicate account translation quota has been exceeded.</summary>
        private const HttpStatusCode HttpStatusCodeQuotaExceeded = (HttpStatusCode)456;

        private const string _serverUrl = "https://api.deepl.com/v2/";

        private const string user_agent = "deepl-dotnet/gever";
        private const string authorisationKey = "DeepL-Auth-Key bb4d9298-9aa1-4f2a-9bdc-6ae0c8b0be74";
        private const string authorisation_header = "Authorization";
        private const string user_agent_header = "User-Agent";

        private readonly HttpClient _httpClient;
        private readonly List<KeyValuePair<string, string>> _headers;

        public DeeplClient()
        {
           
          
            // Create a HttpClientHandler to configure proxy settings
            HttpClientHandler handler = new HttpClientHandler()
            {
                // Create a WebProxy instance with just the proxy address
                Proxy = new WebProxy("{{PROXY_URL}}"),
                // Use proxy
                UseProxy = true
            };

            _httpClient = new HttpClient(handler);


            this._headers = new List<KeyValuePair<string, string>>();
            _headers.Add(new KeyValuePair<string, string>(authorisation_header, authorisationKey));
            _headers.Add(new KeyValuePair<string, string>(user_agent_header, user_agent));
        }

        public HttpResponseMessage ApiPost(string relativeUri, CancellationToken cancellationToken, IEnumerable<(string Key, string Value)> bodyParams = null)
        {
            var requestMessage = new HttpRequestMessage
            {
                RequestUri = new Uri(_serverUrl + relativeUri),
                Method = HttpMethod.Post,
                Content = bodyParams != null
                 ? new LargeFormUrlEncodedContent(
                       bodyParams.Select(pair => new KeyValuePair<string, string>(pair.Key, pair.Value)))
                 : null
            };
            return ApiCall(requestMessage, cancellationToken);
        }



        public HttpResponseMessage ApiUpload(string relativeUri, CancellationToken cancellationToken, IEnumerable<(string Key, string Value)> bodyParams, Stream file, string fileName)
        {
            var content = new MultipartFormDataContent();
            foreach (var (key, value) in bodyParams)
            {
                content.Add(new StringContent(value), key);
            }

            content.Add(new StreamContent(file), "file", fileName);

            var requestMessage = new HttpRequestMessage
            {
                RequestUri = new Uri(_serverUrl + relativeUri),
                Method = HttpMethod.Post,
                Content = content//,
                //Headers = { Accept = { new MediaTypeWithQualityHeaderValue("application/json") } }
            };
            return ApiCall(requestMessage, cancellationToken);
        }

        private HttpResponseMessage ApiCall(HttpRequestMessage requestMessage, CancellationToken cancellationToken)
        {
            try
            {
                foreach (var header in _headers)
                {
                    requestMessage.Headers.Add(header.Key, header.Value);
                }
                return _httpClient.SendAsync(requestMessage, cancellationToken).GetAwaiter().GetResult();
                
            }
            catch (TaskCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException ex)
            {
                throw new DeeplException($"Request timed out: {ex.Message}", ex);
            }
            catch (HttpRequestException ex)
            {
                throw new DeeplException($"Request failed: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new DeeplException($"Unexpected request failure: {ex.Message}", ex);
            }
        }

        public Task<HttpResponseMessage> ApiDeleteAsync(string relativeUri, CancellationToken cancellationToken)
        {
            throw new NotImplementedException();
        }

        public Task<HttpResponseMessage> ApiGetAsync(string relativeUri, CancellationToken cancellationToken, IEnumerable<(string Key, string Value)> queryParams = null, string acceptHeader = null)
        {
            throw new NotImplementedException();
        }

        internal void CheckStatusCode(HttpResponseMessage responseMessage, bool usingGlossary = false, bool downloadingDocument = false)
        {
            var statusCode = responseMessage.StatusCode;
            if (statusCode >= HttpStatusCode.OK && statusCode < HttpStatusCode.BadRequest)
            {
                return;
            }

            string message;
            try
            {

                var errorResult = JsonConvert.DeserializeObject<ErrorResult>(responseMessage.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                message = (errorResult.Message != null ? $", message: {errorResult.Message}" : "") +
                          (errorResult.Detail != null ? $", detail: {errorResult.Detail}" : "");
            }
            catch (JsonException)
            {
                message = string.Empty;
            }

            switch (statusCode)
            {
                case HttpStatusCode.Forbidden:
                    throw new DeeplException("Authorization failure, check AuthKey" + message);
                case HttpStatusCodeQuotaExceeded:
                    throw new DeeplException("Quota for this billing period has been exceeded" + message);
                case HttpStatusCode.NotFound:
                    if (usingGlossary)
                    {
                        throw new DeeplException("Glossary not found" + message);
                    }
                    else
                    {
                        throw new DeeplException("Not found, check ServerUrl" + message);
                    }
                case HttpStatusCode.BadRequest:
                    throw new DeeplException("Bad request" + message);
                case HttpStatusCodeTooManyRequests:
                    throw new DeeplException(
                          "Too many requests, DeepL servers are currently experiencing high load" + message);
                case HttpStatusCode.ServiceUnavailable:
                    if (downloadingDocument)
                    {
                        throw new DeeplException("Document not ready" + message);
                    }
                    else
                    {
                        throw new DeeplException("Service unavailable" + message);
                    }
                default:
                    throw new DeeplException("Unexpected status code: " + statusCode + message);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        protected virtual void Dispose(bool b)
        {
            _httpClient.Dispose();
        }
    }



    internal interface IDeepLClient
    {
        Task<HttpResponseMessage> ApiDeleteAsync(string relativeUri, CancellationToken cancellationToken);
        Task<HttpResponseMessage> ApiGetAsync(string relativeUri, CancellationToken cancellationToken, IEnumerable<(string Key, string Value)> queryParams = null, string acceptHeader = null);
        Task<HttpResponseMessage> ApiPostAsync(string relativeUri, CancellationToken cancellationToken, IEnumerable<(string Key, string Value)> bodyParams = null);
        Task<HttpResponseMessage> ApiUploadAsync(string relativeUri, CancellationToken cancellationToken, IEnumerable<(string Key, string Value)> bodyParams, Stream file, string fileName);
    }

    public class DocumentHandle
    {

        /// <summary>ID of associated document request.</summary>
        public string Document_id { get; }

        /// <summary>Key of associated document request.</summary>
        public string Document_key { get; }
        public DocumentHandle(string document_id, string document_key)
        {
            Document_id = document_id;
            Document_key = document_key;
        }


    }

    public readonly struct ErrorResult
    {
        /// <summary>Initializes a new instance of <see cref="ErrorResult" />, used for JSON deserialization.</summary>

        public ErrorResult(string message, string detail)
        {
            Message = message;
            Detail = detail;
        }

        /// <summary>Message describing the error, if it was included in response.</summary>
        public string Message { get; }

        /// <summary>String explaining more detail the error, if it was included in response.</summary>
        public string Detail { get; }
    }

    public sealed class DocumentStatus
    {
        /*
        public enum StatusCode
        {
            /// <summary>Document translation has not yet started, but will begin soon.</summary>
            [EnumMember(Value = "queued")] Queued,

            /// <summary>Document translation is in progress.</summary>
            [EnumMember(Value = "translating")] Translating,

            /// <summary>Document translation completed successfully, and the translated document may be downloaded.</summary>
            [EnumMember(Value = "done")] Done,

            /// <summary>An error occurred during document translation.</summary>
            [EnumMember(Value = "error")] Error
        }
        */
        /// <summary>Initializes a new <see cref="DocumentStatus" /> object for an in-progress document translation.</summary>
        /// <param name="documentId">Document ID of the associated document.</param>
        /// <param name="status">Status of the document translation.</param>
        /// <param name="secondsRemaining">
        ///   Number of seconds remaining until translation is complete, only included while
        ///   document is in translating state.
        /// </param>
        /// <param name="billedCharacters">
        ///   Number of characters billed for the translation of this document, only included
        ///   after document translation is finished and the state is done.
        /// </param>
        /// <param name="errorMessage">Short description of the error, if available.</param>
        /// <remarks>
        ///   The constructor for this class (and all other Model classes) should not be used by library users. Ideally it
        ///   would be marked <see langword="internal" />, but needs to be <see langword="public" /> for JSON deserialization.
        ///   In future this function may have backwards-incompatible changes.
        /// </remarks>

        public DocumentStatus(string documentId, string status, int secondsRemaining, int billedCharacters, string errorMessage)
        {
            (DocumentId, Status, SecondsRemaining, BilledCharacters, ErrorMessage) =
                  (documentId, status, secondsRemaining, billedCharacters, errorMessage);
        }

        /// <summary>Document ID of the associated document.</summary>
        public string DocumentId { get; }

        /// <summary>Status of the document translation.</summary>
        public string Status { get; }

        /// <summary>
        ///   Number of seconds remaining until translation is complete, only included while
        ///   document is in translating state.
        /// </summary>
        public int SecondsRemaining { get; }

        /// <summary>
        ///   Number of characters billed for the translation of this document, only included
        ///   after document translation is finished and the state is done.
        /// </summary>
        public int BilledCharacters { get; }

        /// <summary>Short description of the error, if available.</summary>
        public string ErrorMessage { get; }

        /// <summary><c>true</c> if no error has occurred during document translation, otherwise <c>false</c>.</summary>
        public bool Ok => Status != "error";

        /// <summary><c>true</c> if document translation has completed successfully, otherwise <c>false</c>.</summary>
        public bool Done => Status == "done";
    }


    public class LargeFormUrlEncodedContent : ByteArrayContent
    {
        private static readonly Encoding Utf8Encoding = Encoding.UTF8;

        public LargeFormUrlEncodedContent(IEnumerable<KeyValuePair<string, string>> nameValueCollection)
              : base(GetContentByteArray(nameValueCollection))
        {
            Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
        }

        private static byte[] GetContentByteArray(IEnumerable<KeyValuePair<string, string>> nameValueCollection)
        {
            if (nameValueCollection == null)
            {
                throw new ArgumentNullException(nameof(nameValueCollection));
            }

            var str = string.Join(
                  "&",
                  nameValueCollection.Select(pair => $"{WebUtility.UrlEncode(pair.Key)}={WebUtility.UrlEncode(pair.Value)}"));
            return Utf8Encoding.GetBytes(str);
        }
    }

    public sealed class TextResult
    {

        public TextResult(string text, string detectedSourceLanguageCode)
        {
            Text = text;
            DetectedSourceLanguageCode = detectedSourceLanguageCode;
        }

        /// <summary>The translated text_param.</summary>
        public string Text { get; }


        public string DetectedSourceLanguageCode { get; }

        /// <summary>Returns the translated text_param.</summary>
        /// <returns>The translated text_param.</returns>
        public override string ToString() => Text;
    }

    readonly struct TextTranslateResult
    {
        public TextResult[] Translations { get; }
        public TextTranslateResult(TextResult[] translations)
        {
            Translations = translations;
        }
    }

    [Serializable]
    public class DeeplException : Exception
    {
        public DeeplException()
        {
        }

        public DeeplException(string message) : base(message)
        {
        }

        public DeeplException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DeeplException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }

    public class DocumentTranslationException : Exception
    {
        public DocumentHandle Handle { get; }

        public DocumentTranslationException()
        {
        }

        public DocumentTranslationException(string message) : base(message)
        {
        }

        public DocumentTranslationException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public DocumentTranslationException(string message, Exception exception, DocumentHandle handle) :
          base(message, exception)
        {
            this.Handle = handle;
        }

        protected DocumentTranslationException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
