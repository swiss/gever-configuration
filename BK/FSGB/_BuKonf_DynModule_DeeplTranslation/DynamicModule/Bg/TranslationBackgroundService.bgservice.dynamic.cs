using ActaNova.Domain;
using ActaNova.Domain.BackgroundService;
using ActaNova.Domain.Extensions;
using ActaNova.Domain.Specialdata.Values;
using ActaNova.Domain.Utilities;
using Microsoft.Scripting.Utils;
using Remotion.Data.DomainObjects;
using Remotion.Data.DomainObjects.Queries;
using Remotion.Logging;
using Remotion.ObjectBinding;
using Remotion.Security;
using Remotion.SecurityManager.Domain.OrganizationalStructure;
using Rubicon.Dms;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Utilities;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Security;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Bukonf.Gever.Domain
{
    public class TranslationBackgroundService : BaseTimeControlledBackgroundService<TranslationBackgroundService>
    {
        private static readonly ILog _logger = LogManager.GetLogger(typeof(TranslationBackgroundService));

        private const string TenantShortName = "DL";
        public int BatchSize { get; set; }


        public TranslationBackgroundService() : base()
        {
            Categories = new string[1] { "WinService_DeeplranslationBackgroundService" };
            GetDisplayName = () => "DeepL Übersetzungsdienst";
            StartHour = 0;
            StopHour = 24;
            Interval = (long)TimeSpan.FromMinutes(1).TotalMilliseconds;
            Actions = new Action[1] { Run };
            BatchSize = 5;
        }

        private void Run()
        {
            _logger.Error("Start Translate BS");
            if (!this.IsStarted) return;

            try
            {
                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (SecurityFreeSection.Activate())
                {
                    var tenant = GetTenantByShortName(TenantShortName);
                    using (TenantSection.SwitchToTenant(Tenant.FindByUnqiueIdentifier(tenant.UniqueIdentifier)))
                    {

                        var queue = FindQueue();
                        var aggList = queue.GetProperty("#TranslationQueueEntry") 
                                                  as FilteredListWithIndex<SpecialdataAggregatePropertyValue>;

                        var openItems = aggList
                                      .Where(e => (string)e.GetProperty("#TranslationQueueEntryStatus") == "open")                            
                                      .Take(BatchSize)
                                      .FetchAllToList()
                                      .Select( o => (SpecialdataAggregatePropertyValue)o);

                        foreach (var item in openItems)
                        {
                            if ((string)item.GetProperty("#TranslationQueueEntryStatus") == "open")
                            {
                                try
                                {
                                    TranslateOne(item);
                                    Remotion.ObjectBinding.BusinessObjectExtensions.SetProperty(item, "#TranslationQueueEntryStatus", "translated");
                                    
                                }
                                catch (Exception e)
                                {
                                    _logger.Error(e.Message + "\n" + e.StackTrace);
                                    Remotion.ObjectBinding.BusinessObjectExtensions.SetProperty(item, "#TranslationQueueEntryStatus", "failed");
                                }
                            }

                        }
                        ClientTransaction.Current.Commit();
                    }
                }

            }
            catch (Exception e)
            {
                _logger.Error(e.Message + "\n" + e.StackTrace);
            }



        }

        private void TranslateOne(SpecialdataAggregatePropertyValue entry)
        {

            var doc = entry.GetProperty("#DocToTranslate") as Document;
            var targetLanguage = (string)entry.GetProperty("#TranslationLanguage");
            _logger.Error(doc);
            _logger.Error(targetLanguage);
            Translate(doc, targetLanguage);
        }

        private void Translate(Document d, string targetLanguage)
        {

            var inputStream = new MemoryStream(d.PrimaryContent.GetContentAsByteArray());
            var outputStream = new MemoryStream(1000);
            var fileName = d.PrimaryContent.Name + "." + d.PrimaryContent.Extension;
            var source_language = ComputeLanguage(d.ContentCulture);
            var target_language = targetLanguage;
            _logger.Error("target_language: " + target_language);
            new DeeplTranslator().TranslateDocument(inputStream, outputStream, fileName, source_language, target_language);

            var translation = Document.NewObject(d.FileHandlingContainer);
            ((IDocument)translation).SetContent(new ByteArrayStreamAccessHandle(outputStream.ToArray()), d.PrimaryContent.Extension, d.PrimaryContent.MimeType);
            translation.PrimaryContent.Name = d.PrimaryContent.Name + "_" + target_language;
        }

        private SpecialdataHostingBaseDataObject FindQueue()
        {

            List<SpecialdataHostingBaseDataObject> fdos = null;
            using (SecurityFreeSection.Activate())
            {
                fdos = QueryFactory.CreateLinqQuery<SpecialdataHostingBaseDataObject>()
                .Where(o => o.Name.Equals("TranslationQueue"))
                .FetchAll()
                .ToList();
            }

            if (fdos.Count == 0)
            {
                _logger.Error("TranslationQueue not found");
            }

            return fdos[0];


        }

        private Tenant GetTenantByShortName(string name)
        {
            return QueryFactory.CreateLinqQuery<Tenant>()
                .Single(t => ((ITenantTrustMixin)t).TenantShortName == name);
        }

        public string ComputeLanguage(string culture)
        {
            switch (culture)
            {
                case ("de"):
                case ("de-CH"):
                    return "de";
                case ("fr"):
                case ("fr-CH"):
                    return "fr";
                case ("it"):
                case ("it-CH"):
                    return "it";
                case ("en"):
                case ("en-CH"):
                    return "en";
                default:
                    return "de-CH";
            }
        }



    }
}

