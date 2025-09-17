// Name: Geschäftsvorfall abschliessen (CloseFileCase))
// Version: 2.0.1
// Datum: 26.06.2024
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
  public class CloseFileCaseActivityCommandModule : GeverActivityCommandModule
  {
    public class CloseFileCaseParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (FileCase))]
      public TypedParsedExpression Target { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression SkipIfAlreadyClosed { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression SkipIfCanceled { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Geschäftsvorfall kann nicht abgeschlossen werden. Gründe: {1}")]
      CanNotCloseError
    }

    public CloseFileCaseActivityCommandModule ()
        : base ("CloseFileCase:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<CloseFileCaseParameters> (commandActivity);
      var targetExpression = parameters.Target;
      var skipIfClosedExpression = parameters.SkipIfAlreadyClosed;
      var shouldSkipIfClosed = skipIfClosedExpression?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      var fileCase = targetExpression?.Invoke<FileCase> (commandActivity, commandActivity.WorkItem);
      var skipIfCanceledExpression = parameters.SkipIfCanceled;
      var shouldSkipIfCanceled = skipIfCanceledExpression?.Invoke<bool> (commandActivity, commandActivity.WorkItem);

      if (fileCase.FileHandlingState == FileHandlingStateType.Canceled && shouldSkipIfCanceled == true)
        return true;

      if (fileCase == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(CloseFileCaseParameters.Target)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(CloseFileCaseParameters.Target)));
      }

      if (fileCase.CanCloseObject())
      {
        fileCase.CloseObject();
      }
      else if (!(fileCase.FileHandlingState == FileHandlingStateType.Closed && shouldSkipIfClosed == true))
      {
        var operation = FileHandlingObjectStateMachineDefinition.Get<FileHandlingObjectStateMachineDefinition.CloseOperation>().For (fileCase, default(object));
        operation.IsAvailable (out var reasons);


        throw new ActivityCommandException (
                $"Can not close file case in activity '{commandActivity.ID}'. " + reasons)
            .WithUserMessage (ModuleLocalization.CanNotCloseError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, reasons));
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<CloseFileCaseParameters> (commandActivity);
        var targetExpression = parameters.Target;

        if (targetExpression == null)
        {
          throw new ActivityCommandException (
                  $"Parameter '{nameof(CloseFileCaseParameters.Target)}' must not be null in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, nameof(CloseFileCaseParameters.Target)));
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