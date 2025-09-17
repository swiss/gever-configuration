/*
* ===================================================================
*(c)Lufthansa Industry Solutions GmbH 2023, CH, All rights reserved
* SendMailCommandModule
* @author       dimitrij.zaks@lhind.dlh.de
* @version      V1.2
* @date         30.09.2023
* @description  This module sets language on target document(s).
* =================================================================
*/
using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using Aspose.Pdf.Operators;
using Remotion.Data.DomainObjects.ObjectBinding;
using Remotion.Globalization;
using Remotion.Logging;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Bukonf.Gever.Domain
{
    [LocalizationEnum]
    public enum LocalizedUserMessagesSetDocumentLanguage
    {
        [MultiLingualName("Das Dokument konnte nicht gefunden werden.", ""),
         MultiLingualName("Das Dokument konnte nicht gefunden werden.", "De"),
         MultiLingualName("Le document est introuvable.", "Fr"),
         MultiLingualName("Impossibile trovare il documento.", "It")]
        NoDocument,

        [MultiLingualName("Die Dokumente konnten nicht gefunden werden.", ""),
         MultiLingualName("Die Dokumente konnten nicht gefunden werden.", "De"),
         MultiLingualName("Les documents sont introuvables.", "Fr"),
         MultiLingualName("Impossibile trovare trovare i documenti.", "It")]
        NoDocuments,

        [MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", ""),
         MultiLingualName("Die angegebene Sprache \"{0}\" wurde nicht gefunden.", "De"),
         MultiLingualName("La langue spécifiée \"{0}\" est introuvable.", "Fr"),
         MultiLingualName("Impossibile trovare la lingua specificata \"{0}\".", "It")]
        UnknownLanguage,

        [MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden.", ""),
         MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden.", "De"),
         MultiLingualName("L’objet métier \"{0}\" de l’activité ne peut être édité.", "Fr"),
         MultiLingualName("L'oggetto business \"{0}\" dell'attività non può essere modificato.", "It")]
        NotEditable
    }

    public class SetDocumentLanguageCommandModule : GeverActivityCommandModule
    {
        private static readonly ILog _logger = LogManager.GetLogger(typeof(SetDocumentLanguageCommandModule));
        private static readonly string[] Cultures = new string[] { "de", "de-CH", "fr", "fr-CH", "it", "it-CH", "en", "en-CH" };
     

        public class SetDocumentLanguageInputParameters : GeverActivityCommandParameters
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
            public TypedParsedExpression Language { get; set; }

        }

        public SetDocumentLanguageCommandModule() : base(
            "SetDocumentLanguage:ActivityCommandClassificationType")
        {        
        }

        public override bool TryExecute(CommandActivity commandActivity)
        {     
            var parameters  = ActivityCommandParser.Parse<SetDocumentLanguageInputParameters>(commandActivity);   
            var targetDocuments = GetTargetDocuments(parameters, commandActivity);
            targetDocuments.ForEach(t => SetLangauge(t, parameters, commandActivity));
            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            try
            {
                var parameters = ActivityCommandParser.Parse<SetDocumentLanguageInputParameters>(commandActivity);
                new ParameterValidator(commandActivity).ValidateParameters(parameters);
            }
            catch (Exception ex)
            {
                _logger.Error($"Validation of CommandActivity '{commandActivity.ID}' failed.", ex);
                return ex.GetUserMessage() ?? ex.Message;
            }

            return null;
        }

        private void SetLangauge(Document d, SetDocumentLanguageInputParameters parameters, CommandActivity commandActivity)
        {
            if(!d.CanEdit())
            {
                throw new ActivityCommandException($"User cannot edit {d}")
                                .WithUserMessage(LocalizedUserMessagesSetDocumentLanguage.NotEditable
                                    .ToLocalizedName().FormatWith(d));
            }
            var culture = parameters.Language.Invoke<string>((object)commandActivity, (object)commandActivity.WorkItem);
            d.ContentCulture = ValidateCulture(culture);

        }

        private IEnumerable<Document> GetTargetDocuments(SetDocumentLanguageInputParameters parameters, CommandActivity commandActivity)
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
            if (result == null || ( result.GetType().Equals(typeof(IEnumerable<Document>)) && !((IEnumerable<Document>)result).Any() ))
            {
                throw new ActivityCommandException($"Expression hat keine Ergebnisse geliefert'")
                      .WithUserMessage(LocalizedUserMessagesSetDocumentLanguage.NoDocument.ToLocalizedName());
            }
        }   

        private string ValidateCulture(string culture)
        {
            if(!Cultures.Contains(culture))
            {
                throw new ActivityCommandException($"Invalid Culture '{culture}'.")
                        .WithUserMessage(LocalizedUserMessagesSetDocumentLanguage.UnknownLanguage.ToLocalizedName()
                            .FormatWith(culture));
            }
            return ComputeCultue(culture);
            
         }

        public string ComputeCultue(string culture)
        {
            switch (culture)
            {
                case ("de"):
                case ("de-CH"):
                    return "de-CH";
                case ("fr"):
                case ("fr-CH"):
                    return "fr-CH";
                case ("it"):
                case ("it-CH"):
                    return "it-CH";
                case ("en"):
                case ("en-CH"):
                    return "en-CH";
                default:
                    return "de-CH";
            }
        }

     
    }

}