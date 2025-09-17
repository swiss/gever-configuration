// Name: Set Activity Flow Variable Value
// Version: 3.0.0
// Datum: 05.12.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu; Ursprünglich LHIND - philipp.roessler@gs-ejpd.admin.ch

// Mit dieser Befehlsaktivität werden die im Parameter übergebenen FlowVariablen gesetzt.
// Wenn eine übergebene Variable nicht exisitiert, wird sie erstellt.
// Version 1.0 - farid.modarressi@lhind.dlh.de - 04.02.2021
// Version 1.1 - philipp.roessler@gs-ejpd.admin.ch - 04.05.2021

using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Utilities;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Utilities.Expr;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using System;
using Remotion.Data.DomainObjects;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace LHIND.BBL.Beschaffung
{
  public enum LocalizedUserMessages
  {
    [De ("Befehlsaktivität \"{0}\": Der Ausdruck kann nicht verarbeitet werden \"{1}\"."),
     MultiLingualName ("Command activity \"{0}\": The expression cannot be processed \"{1}\".", "En"),
     MultiLingualName ("Activité de commande \"{0}\": L'expression ne peut pas être traitée \"{1}\".", "Fr"),
     MultiLingualName ("Attività di comando \"{0}\": L'espressione non può essere elaborata \"{1}\".", "It")]
    InvalidExpression,

    [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" hat keinen gültigen Typ.")]
    InvalidParameterType,

    [De ("Befehlsaktivität \"{0}\": Die Aktivität muss schreibgeschützt sein.")]
    InstanceNotReadOnlyError,
  }

  public class SetActivityFlowVariableValueCommandModule : GeverActivityCommandModule, ITemplateOnlyActivityCommandModule
  {
    public class SetActivityFlowVariableValueParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression FlowVariableName { get; set; }

      [ExpectedExpressionType (typeof (object))]
      public TypedParsedExpression FlowVariableValue { get; set; }
    }

    public SetActivityFlowVariableValueCommandModule () : base (
        "LHIND_SetActivityFlowVariableValue:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<SetActivityFlowVariableValueParameters> (commandActivity);
      var flowVariableNameExpression = parameters.FlowVariableName;
      var flowVariableValueExpression = parameters.FlowVariableValue;

      var flowVariableName = flowVariableNameExpression.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var flowVariableValue = flowVariableValueExpression.Invoke<object> (commandActivity, commandActivity.WorkItem);

      var flowVariableObject = commandActivity.GetOrCreateFlowVariable (flowVariableName);

      switch (flowVariableValue)
      {
        case bool value:
          flowVariableObject.SetValue (value);
          break;
        case int value:
          flowVariableObject.SetValue (value);
          break;
        case double value:
          flowVariableObject.SetValue (value);
          break;
        case string value:
          flowVariableObject.SetValue (value);
          break;
        case ObjectID value:
          flowVariableObject.SetValue (value);
          break;
        default:
          if (flowVariableValue.GetType().IsSubclassOf (typeof (DomainObject)))
            flowVariableObject.SetValue ((DomainObject)flowVariableValue);
          else
          {
            throw new ActivityCommandException ($"The parameter does not have a valid type in activity '{commandActivity.ID}'.")
                .WithUserMessage (LocalizedUserMessages.InvalidParameterType.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, nameof(flowVariableValue)));
          }

          break;
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<SetActivityFlowVariableValueParameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);

        if (!commandActivity.InstanceReadOnly)
        {
          throw new ActivityCommandException ($"Activity must be set to instance read only in activity '{commandActivity.ID}'.")
              .WithUserMessage (LocalizedUserMessages.InstanceNotReadOnlyError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
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