/*
 * Name: SetSedexRecipient
 * Version: 2.0.0
 * Datum: 10.06.2024
 * Autor: Lufthansa Industry Solution GmbH – Dimitrij Zaks - dimitrij.zaks@lhind.dlh.de | Für BS4.0 RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu
 * Beschreibung: Die Befehlsaktivität SetSedexRecipient hat zum Ziel die Empfänger von SEDEX-Paketen während der Ausführung (dynamisch) zu ermitteln.
 */

using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Sedex.Domain;
using Rubicon.Gever.Bund.Sedex.Domain.Outbox;
using Rubicon.Gever.Bund.Sedex.Domain.Outbox.Data;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using System;
using System.Linq;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;

namespace Rubicon.Gever.Bund.Domain.Workflow
{
  public class SetSedexRecipientActivityCommandModule : GeverActivityCommandModule
  {
    public enum LocalizationSetSedexRecipient
    {
      [De ("Befehlsaktivität \"{0}\": Die Nachfolgeaktivität ist im Prozess nicht eindeutig oder nicht vorhanden. Passen Sie den Prozess an.")]
      NextActivityError,

      [De ("Befehlsaktivität \"{0}\": Der Wert des Parameters \"{0}\" konnte nicht gefunden werden.")]
      ParameterParsingResultEmpty
    }

    public class SetSedexRecipientInputParameters : GeverActivityCommandParameters
    {
      [RequiredOneOfMany (new[] { nameof(SedexParticipant), nameof(SedexDistributionList) })]
      [EitherOr (new[] { nameof(SedexParticipant), nameof(SedexDistributionList) })]
      [ExpectedExpressionType (typeof (SedexParticipant))]
      public TypedParsedExpression SedexParticipant { get; set; }

      [RequiredOneOfMany (new[] { nameof(SedexParticipant), nameof(SedexDistributionList) })]
      [EitherOr (new[] { nameof(SedexParticipant), nameof(SedexDistributionList) })]
      [ExpectedExpressionType (typeof (SedexDistributionList))]
      public TypedParsedExpression SedexDistributionList { get; set; }

      [ExpectedExpressionType (typeof (SedexDataWriterClassificationType))]
      public TypedParsedExpression SedexDataWriterClassificationType { get; set; }
    }

    public SetSedexRecipientActivityCommandModule ()
        : base ("SetSedexRecipient:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var next = GetNextActivity (commandActivity);
      var parameters = ActivityCommandParser.Parse<SetSedexRecipientInputParameters> (commandActivity);
      next.SedexRecipient = ResolveParameterValue<SedexParticipant> (commandActivity, parameters.SedexParticipant) ?? next.SedexRecipient;
      next.SedexDistributionList =
          ResolveParameterValue<SedexDistributionList> (commandActivity, parameters.SedexDistributionList) ?? next.SedexDistributionList;
      next.DataWriterType = ResolveParameterValue<SedexDataWriterClassificationType> (commandActivity, parameters.SedexDataWriterClassificationType)
                            ?? next.DataWriterType;
      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<SetSedexRecipientInputParameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);
        GetNextActivity (commandActivity);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    private T ResolveParameterValue<T> (CommandActivity commandActivity, TypedParsedExpression expr)
    {
      if (expr == null)
        return default(T);

      var result = (T)expr.Invoke (commandActivity, commandActivity.WorkItem);

      if (result == null)
        throw new ActivityCommandException ($"Error while parsing '{expr}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                LocalizationSetSedexRecipient.ParameterParsingResultEmpty.ToLocalizedName().FormatWith (commandActivity.DisplayName, nameof(expr)));

      return result;
    }

    private SedexSendActivity GetNextActivity (CommandActivity commandActivity)
    {
      var nextActivitiesAfter =
          commandActivity.GetNextActivitiesAfter (new ActivityBase.GetNextActivitiesSpecification<ActivityBase>());
      return nextActivitiesAfter.Count == 1 && nextActivitiesAfter[0] is SedexSendActivity sedexActivity
          ? sedexActivity
          : throw new ActivityCommandException (
                  $"Activity after '{commandActivity.ID}' with ID '{nextActivitiesAfter.FirstOrDefault()?.ID}' is not of type SedexSendActivity.")
              .WithUserMessage (LocalizationSetSedexRecipient.NextActivityError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
    }
  }
}