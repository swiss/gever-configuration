using Newtonsoft.Json;
using Remotion.Logging;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;


namespace Bukonf.Gever.Domain
{
    public class DeeplTranslator
    {
        private static readonly ILog _logger = LogManager.GetLogger(typeof(DeeplTranslator));

        private const string target_lang_param = "target_lang";
        private const string source_lang_param = "source_lang";
        private const string text_param = "text";

    
        private readonly DeeplClient _client;

        public DeeplTranslator()
        {          
            _client = new DeeplClient();
        }


        public string Translate(string txt, string src, string trg)
        {
            _logger.Debug("Start translating");              
            var bodyParams = new (string Key, string Value)[] {
             (source_lang_param, src),
             (target_lang_param, trg),
             (text_param, txt) 
            };

            var response = _client.ApiPost("translate", CancellationToken.None, bodyParams);
            var resultRoh = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            var result = JsonConvert.DeserializeObject<TextTranslateResult>(resultRoh);
            var translation = result.Translations[0];
            _logger.Debug("End translating");
            return translation.Text;
        }

        public bool TranslateDocument(
                    FileInfo inputFileInfo,
                    FileInfo outputFileInfo,
                    string sourceLanguageCode,
                    string targetLanguageCode,
                    CancellationToken cancellationToken = default)
        {
            using (var inputFile = inputFileInfo.OpenRead())
            using (var outputFile = outputFileInfo.Open(FileMode.CreateNew, FileAccess.Write))
            {
                try
                {
                    TranslateDocument(
                          inputFile,
                          outputFile,
                          inputFileInfo.Name,
                          sourceLanguageCode,
                          targetLanguageCode,
                          cancellationToken);
                }
                catch(Exception e)
                {
                    try
                    {
                        outputFileInfo.Delete();
                    }
                    catch
                    {
                        // ignored
                    }
                    _logger.Error(e);
                    throw;
                }

            }
            return true;  
        }

        public void TranslateDocument(Stream inputFile, Stream outputFile, string fileName, string src, string trg, CancellationToken cancellationToken = default)
        {
            _logger.Debug("Start translating document");
            DocumentHandle handle = null;
            try
            {
                handle = UploadDocument(
                            inputFile,
                            fileName,
                            src,
                            trg,
                            cancellationToken);
                _logger.Debug("Document uploaded");
                TranslateDocumentWaitUntilDoneAsync(handle, cancellationToken);
                TranslateDocumentDownload(handle, outputFile, cancellationToken);
                _logger.Debug("End translating document");
            }
            catch (Exception exception)
            {
                _logger.Error(exception);
                throw new DocumentTranslationException(
                      $"Error occurred during document translation: {exception.Message}",
                      exception,
                      handle);
            }
        }

        public DocumentHandle UploadDocument(Stream file, string fileName, string src, string trg, CancellationToken cancellationToken = default)
        {
            var bodyParams = new (string Key, string Value)[] {
                (source_lang_param, src),
                (target_lang_param, trg)
            };

            var response = _client.ApiUpload("document", cancellationToken, bodyParams, file, fileName);
            _client.CheckStatusCode(response);
            var resultRoh = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            var result = JsonConvert.DeserializeObject<DocumentHandle>(resultRoh);
            return result;
        }

        public void TranslateDocumentWaitUntilDoneAsync(
         DocumentHandle handle,
         CancellationToken cancellationToken = default)
        {
            var status = TranslateDocumentStatus(handle, cancellationToken);
            while (status.Ok && !status.Done)
            {
                //await Task.Delay(CalculateDocumentWaitTime(), cancellationToken).ConfigureAwait(false);
                Task.Delay(CalculateDocumentWaitTime()).Wait(); // Wait 2 seconds with blocking
                status = TranslateDocumentStatus(handle, cancellationToken);
            }

            if (!status.Ok)
            {
                throw new DeeplException(status.ErrorMessage ?? "Unknown error");
            }
        }
        public DocumentStatus TranslateDocumentStatus(DocumentHandle handle,CancellationToken cancellationToken = default)
        {
            var bodyParams = new (string Key, string Value)[] { ("document_key", handle.Document_key) };
            var responseMessage =
                   _client.ApiPost(
                              $"document/{handle.Document_id}",
                              cancellationToken,
                              bodyParams);
            _client.CheckStatusCode(responseMessage);
            return  JsonConvert.DeserializeObject<DocumentStatus>(responseMessage.Content.ReadAsStringAsync().GetAwaiter().GetResult());
        }

        public void TranslateDocumentDownload(DocumentHandle handle, FileInfo outputFileInfo, CancellationToken cancellationToken = default)
        {
            using (var outputFileStream = outputFileInfo.Open(FileMode.CreateNew, FileAccess.Write))
            {
                try
                {
                    TranslateDocumentDownload(handle, outputFileStream, cancellationToken);
                }
                catch (Exception e)
                {
                    try
                    {
                        outputFileInfo.Delete();
                    }
                    catch
                    {
                        // ignored
                    }
                    _logger.Error(e);
                    throw;
                }
            }
           
        }

        public void TranslateDocumentDownload(DocumentHandle handle, Stream outputFile, CancellationToken cancellationToken = default)
        {
            var bodyParams = new (string Key, string Value)[] { ("document_key", handle.Document_key) };
            using (var responseMessage = _client.ApiPost( $"document/{handle.Document_id}/result",  cancellationToken,bodyParams))
            {
                _client.CheckStatusCode(responseMessage);
                responseMessage.Content.CopyToAsync(outputFile).ConfigureAwait(false).GetAwaiter().GetResult();
            }        
        }


        private static TimeSpan CalculateDocumentWaitTime()
        {
            // hintSecondsRemaining is currently unreliable, so just poll equidistantly
            const int POLLING_TIME_SECS = 5;
            return TimeSpan.FromSeconds(POLLING_TIME_SECS);
        }

    }


}
