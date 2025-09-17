
using ActaNova.Domain;
using ActaNova.Domain.Specialdata;
using ActaNova.Domain.Workflow;
using Remotion.Data.DomainObjects.Queries;
using Remotion.Globalization;
using Remotion.Logging;
using Remotion.Security;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Bukonf.Gever.Domain
{
    [LocalizationEnum]
    public enum LocalizedUserMessagesTranslateDocumentAsync
    {
        [MultiLingualName("Das Dokument konnte nicht gefunden werden.", ""),
         MultiLingualName("Das Dokument konnte nicht gefunden werden.", "De"),
         MultiLingualName("Das Dokument konnte nicht gefunden werden.", "Fr"),
         MultiLingualName("Das Dokument konnte nicht gefunden werden.", "It")]
        NoDocument,

        [MultiLingualName("Die Dokumente konnten nicht gefunden werden.", ""),
         MultiLingualName("Die Dokumente konnten nicht gefunden werden.", "De"),
         MultiLingualName("Die Dokumente konnten nicht gefunden werden.", "Fr"),
         MultiLingualName("Die Dokumente konnten nicht gefunden werden.", "It")]
        NoDocuments,

        [MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", ""),
         MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", "De"),
         MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", "Fr"),
         MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", "It")]
        UnknownLanguage,

        [MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", ""),
         MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", "De"),
         MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", "Fr"),
         MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", "It")]
        QueueNotFound,

        [MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden.", ""),
         MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden.", "De"),
         MultiLingualName("L’objet métier \"{0}\" de l’activité ne peut être édité.", "Fr"),
         MultiLingualName("L'oggetto business \"{0}\" dell'attività non può essere modificato.", "It")]
        NotEditable
    }

    public class TranslateDocumentAsyncCommandModule : GeverActivityCommandModule
    {
        private static readonly ILog _logger = LogManager.GetLogger(typeof(TranslateDocumentAsyncCommandModule));

        public class TranslateDocumentSyncInputParameters : GeverActivityCommandParameters
        {
            [ExpectedExpressionType(typeof(Document))]
            [RequiredOneOfMany(new string[] { "Target", "Targets" })]
            [EitherOr(new string[] { "Target", "Targets" })]
            public TypedParsedExpression Target { get; set; }

            [ExpectedExpressionType(typeof(IEnumerable<Document>))]
            [RequiredOneOfMany(new string[] { "Target", "Targets" })]
            [EitherOr(new string[] { "Target", "Targets" })]
            public TypedParsedExpression Targets { get; set; }

            [ExpectedExpressionType(typeof(string))]
            [Required]
            public string Language { get; set; }

        }

        public TranslateDocumentAsyncCommandModule() : base(
            "TranslateDocumentAsync:ActivityCommandClassificationType")
        {
        }

        public override bool TryExecute(CommandActivity commandActivity)
        {
            var parameters = ActivityCommandParser.Parse<TranslateDocumentSyncInputParameters>(commandActivity);
            var targetDocuments = GetTargetDocuments(parameters, commandActivity);
            targetDocuments.ForEach(t => Translate(t, parameters, commandActivity));
            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            try
            {
                var parameters = ActivityCommandParser.Parse<TranslateDocumentSyncInputParameters>(commandActivity);
                new ParameterValidator(commandActivity).ValidateParameters(parameters);
            }
            catch (Exception ex)
            {
                _logger.Error($"Validation of CommandActivity '{commandActivity.ID}' failed.", ex);
                return ex.GetUserMessage() ?? ex.Message;
            }

            return null;
        }

        private void Translate(Document d, TranslateDocumentSyncInputParameters parameters, CommandActivity commandActivity)
        {

            var queue = FindQueue(commandActivity);
            var newValue = BaseObjectSpecialdataAggregateExtensions.AddAggregateListValue(queue, "#TranslationQueueEntry");
            Remotion.ObjectBinding.BusinessObjectExtensions.SetProperty(newValue, "#DocToTranslate", d);
            Remotion.ObjectBinding.BusinessObjectExtensions.SetProperty(newValue, "#TranslationLanguage", parameters.Language);
            Remotion.ObjectBinding.BusinessObjectExtensions.SetProperty(newValue, "#TranslationQueueEntryStatus", "open");
        }

        private SpecialdataHostingBaseDataObject FindQueue(CommandActivity commandActivity)
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
                Throw(commandActivity, LocalizedUserMessagesTranslateDocumentAsync.QueueNotFound);
            }

            return fdos[0];

        }

        private void Throw(CommandActivity commandActivity, LocalizedUserMessagesTranslateDocumentAsync msg)
        {
            throw new ActivityCommandException($"Error while executing activity '{commandActivity.ID}'.")
               .WithUserMessage(msg.ToLocalizedName().FormatWith((object)commandActivity.DisplayName));
        }

        private IEnumerable<Document> GetTargetDocuments(TranslateDocumentSyncInputParameters parameters, CommandActivity commandActivity)
        {
            return parameters.Target != null ? CreateTarget(parameters.Target, commandActivity) : CreateTargets(parameters.Targets, commandActivity);
        }

        private IEnumerable<Document> CreateTarget(TypedParsedExpression target, CommandActivity commandActivity)
        {
            var dok = target.Invoke<Document>((object)commandActivity, (object)commandActivity.WorkItem);
            CheckResult(dok);
            return new List<Document>() { dok };
        }

        private IEnumerable<Document> CreateTargets(TypedParsedExpression targets, CommandActivity commandActivity)
        {
            var doks = targets.Invoke<IEnumerable<Document>>((object)commandActivity, (object)commandActivity.WorkItem);
            CheckResult(doks);
            return doks;
        }

        private static void CheckResult(Object result)
        {
            if (result == null || (result.GetType().Equals(typeof(IEnumerable<Document>)) && !((IEnumerable<Document>)result).Any()))
            {
                throw new ActivityCommandException($"Expression hat keine Ergebnisse geliefert'")
                      .WithUserMessage(LocalizedUserMessagesTranslateDocumentAsync.NoDocument.ToLocalizedName());
            }
        }
    }

}