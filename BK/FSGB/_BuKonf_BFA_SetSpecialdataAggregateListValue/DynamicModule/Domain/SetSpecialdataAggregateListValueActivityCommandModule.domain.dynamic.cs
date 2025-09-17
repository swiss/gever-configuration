// Name: Fachdaten-Aggregat manipulieren (SetSpecialdataAggregateListValue)
// Version: 1.0.1
// Datum: 22.04.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu

using System;
using System.Collections.Generic;
using System.Linq;
using ActaNova.Domain.Specialdata;
using ActaNova.Domain.Specialdata.Properties.Calculated;
using ActaNova.Domain.Specialdata.Values;
using ActaNova.Domain.Workflow;
using Newtonsoft.Json;
using Remotion.Data.DomainObjects;
using Remotion.Globalization;
using Remotion.ObjectBinding;
using Rubicon.Domain;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  public class SetSpecialdataAggregateListValueActivityCommandModule : GeverActivityCommandModule
  {
    public class ModuleParameters : GeverActivityCommandParameters
    {
      public string Values { get; set; }

      [ExpectedExpressionType (typeof (int))]
      public TypedParsedExpression Position { get; set; }

      [Required]
      public string Mode { get; set; }

      [Required]
      [ExpectedExpressionType (typeof (BaseTenantBoundObject))]
      public TypedParsedExpression HostObject { get; set; }

      [Required]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression PropertyIdentifier { get; set; }
    }

    private enum Modes
    {
      Add,
      Edit,
      Delete
    };

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt \"{1}\" kann nicht bearbeitet werden.")]
      TargetCannotEditError,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt \"{1}\" kann nicht entfernt werden.")]
      TargetCannotDeleteError,

      [De ("Befehlsaktivität \"{0}\": Der Modus \"{1}\" wird nicht unterstützt. Verfügbare Modi: \"{2}\"")]
      UnsupportedModeError,

      [De ("Befehlsaktivität \"{0}\": Das Aggregat hat kein Element an Position \"{1}\"")]
      InvalidPositionError,

      [De ("Befehlsaktivität \"{0}\": Syntaxfehler. Geben Sie den Parameter mit korrekter Abfragesyntax an. Erwarteter Ergebnistyp: \"{1}\"")]
      ExpressionParseError,
    }

    public SetSpecialdataAggregateListValueActivityCommandModule ()
        : base ("SetSpecialdataAggregateListValue:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<ModuleParameters> (commandActivity);

      var hostObject = parameters.HostObject?.Invoke<BaseTenantBoundObject> (commandActivity, commandActivity.WorkItem);
      CheckParameterNotNull (commandActivity, hostObject, nameof(ModuleParameters.HostObject));

      var propertyIdentifier = parameters.PropertyIdentifier?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      CheckParameterNotNull (commandActivity, propertyIdentifier, nameof(ModuleParameters.PropertyIdentifier));

      var listProperty = hostObject.GetProperty (propertyIdentifier) as IEnumerable<SpecialdataAggregatePropertyValue>;

      var modeSelected = Enum.TryParse (parameters.Mode, true, out Modes mode);

      if (!modeSelected)
        ThrowUnsupportedModeException (commandActivity, parameters.Mode);

      dynamic aggregateValues = null;

      if (mode != Modes.Delete)
        aggregateValues = JsonConvert.DeserializeObject (parameters.Values);

      var position = -1;

      if (mode != Modes.Add)
      {
        var invokedPosition = parameters.Position?.Invoke<int> (commandActivity, commandActivity.WorkItem);
        CheckParameterNotNull (commandActivity, invokedPosition, nameof(ModuleParameters.Position));

        position = invokedPosition ?? -1;
      }

      switch (mode)
      {
        case Modes.Add:
          AddEntry (hostObject, propertyIdentifier, aggregateValues, commandActivity);
          break;
        case Modes.Edit:
          EditEntry (GetPropertyAtPosition (listProperty, position, commandActivity), aggregateValues, commandActivity);
          break;
        case Modes.Delete:
          DeleteEntry (GetPropertyAtPosition (listProperty, position, commandActivity), commandActivity);
          break;
        default:
          ThrowUnsupportedModeException (commandActivity, parameters.Mode);
          break;
      }

      var service = ClientTransactionScope.CurrentTransaction.RootTransaction.GetTransactionBoundService<SpecialdataCalculatedPropertyValueCacheService>();
      service.InvalidateCache();

      return true;
    }

    private void CheckParameterNotNull (CommandActivity commandActivity, object invokedObject, string parameterName)
    {
      if (invokedObject == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{parameterName}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.ParameterNullError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, parameterName));
      }
    }

    private void AddEntry (BaseTenantBoundObject hostObject, string propertyIdentifier, dynamic aggregateValues, CommandActivity commandActivity)
    {
      if (!hostObject.CanEdit (ignoreSuspendable: true))
      {
        throw new ActivityCommandException (
                $"User cannot Edit '{hostObject.ID}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.TargetCannotEditError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    hostObject?.DisplayName));
      }

      var newValue = hostObject.AddAggregateListValue (propertyIdentifier);

      SetAggregateListValueProperty (aggregateValues, newValue, commandActivity);
    }

    private void EditEntry (SpecialdataAggregatePropertyValue listEntry, dynamic aggregateValues, CommandActivity commandActivity)
    {
      if (!listEntry.CanEdit (ignoreSuspendable: true))
      {
        throw new ActivityCommandException (
                $"User cannot Edit '{listEntry.ID}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.TargetCannotEditError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    listEntry?.DisplayName));
      }

      SetAggregateListValueProperty (aggregateValues, listEntry, commandActivity);
    }

    private void DeleteEntry (SpecialdataAggregatePropertyValue listEntry, CommandActivity commandActivity)
    {
      if (!listEntry.CanDelete())
      {
        throw new ActivityCommandException (
                $"User cannot Delete '{listEntry.ID}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.TargetCannotDeleteError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    listEntry?.DisplayName));
      }

      listEntry.Delete();
    }

    private void SetAggregateListValueProperty (dynamic aggregateValues, SpecialdataAggregatePropertyValue listValue, CommandActivity commandActivity)
    {
      foreach (var property in aggregateValues)
      {
        string identifier = property.identifier;
        string value = property.value.Value;

        var actualValue = GetValueFromExpression (commandActivity, value);

        listValue.SetProperty (identifier, actualValue);
      }
    }

    private object GetValueFromExpression (CommandActivity commandActivity, string expressionString)
    {
      var workItemType = commandActivity.WorkItem?.GetPublicDomainObjectType() ?? commandActivity.EffectiveWorkItemType;
      var formattedExpressionString = $"(activity, workItem) => {expressionString}";

      var parsedExpression = ExpressionConfiguration.Default.Parse (formattedExpressionString, typeof (CommandActivity), workItemType)
          .WithExpectedReturnType (typeof (object));

      var parseException = parsedExpression.Exception;
      if (!parsedExpression.IsValid || parseException != null)
      {
        throw new ActivityCommandException ($"Error while parsing expression '{expressionString}' of activity '{commandActivity.ID}'.", parseException)
            .WithUserMessage (ModuleLocalization.ExpressionParseError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(Object)));
      }

      var actualObject = parsedExpression?.Invoke<object> (commandActivity, commandActivity.WorkItem);
      return actualObject;
    }

    private SpecialdataAggregatePropertyValue GetPropertyAtPosition (
        IEnumerable<SpecialdataAggregatePropertyValue> aggregateEntries,
        int position,
        CommandActivity commandActivity)
    {
      var property = aggregateEntries.ElementAtOrDefault (position);

      if (property == null)
      {
        throw new ActivityCommandException (
                $"Aggregate position {position} does not exist in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.InvalidPositionError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, position));
      }

      return property;
    }

    private void ThrowUnsupportedModeException (CommandActivity commandActivity, string mode)
    {
      throw new ActivityCommandException (
              $"Unsupported mode '{mode}' in activity '{commandActivity.ID}'.")
          .WithUserMessage (
              ModuleLocalization.UnsupportedModeError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, mode, string.Join (", ", Enum.GetNames (typeof (Modes)))));
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<ModuleParameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }
  }
}