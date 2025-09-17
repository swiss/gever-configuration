// Name: Geschäftsvorfall auslösen (StartFileCase)
// Version: 2.0.0
// Datum: 04.07.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu

using System;
using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  public class StartFileCaseActivityCommandModule : GeverActivityCommandModule
  {
    public class StartFileCaseParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (FileCase))]
      public TypedParsedExpression Target { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression SkipIfAlreadyStarted { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Geschäftsvorfall kann nicht ausgelöst werden. Gründe: {1}")]
      CanNotStartError
    }

    public StartFileCaseActivityCommandModule ()
        : base ("StartFileCase:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<StartFileCaseParameters> (commandActivity);
      var targetExpression = parameters.Target;
      var skipExpression = parameters.SkipIfAlreadyStarted;

      var shouldSkip = skipExpression?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      var fileCase = targetExpression?.Invoke<FileCase> (commandActivity, commandActivity.WorkItem);

      if (fileCase == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(StartFileCaseParameters.Target)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(StartFileCaseParameters.Target)));
      }

      if (fileCase.CanStartObject())
      {
        fileCase.StartObject();
      }
      else if (!(fileCase.FileHandlingState == FileHandlingStateType.Work && shouldSkip == true))
      {
        var operation = FileHandlingObjectStateMachineDefinition.Get<FileHandlingObjectStateMachineDefinition.StartOperation>().For (fileCase, default(object));
        operation.IsAvailable (out var reasons);

        throw new ActivityCommandException (
                $"Can not start file case in activity '{commandActivity.ID}'. " + reasons)
            .WithUserMessage (ModuleLocalization.CanNotStartError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, reasons));
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<StartFileCaseParameters> (commandActivity);
        var targetExpression = parameters.Target;

        if (targetExpression == null)
        {
          throw new ActivityCommandException (
                  $"Parameter '{nameof(StartFileCaseParameters.Target)}' must not be null in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, nameof(StartFileCaseParameters.Target)));
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