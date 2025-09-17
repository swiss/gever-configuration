// Name: Arbeitsanweisung setzen / Bemerkung setzen (setRemark)
// Version: 2.0.0
// Datum: 04.07.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu

using System;
using ActaNova.Domain;
using ActaNova.Domain.DocumentFields.Definition;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  public class SetRemarkActivityCommandModule : GeverActivityCommandModule
  {
    public class RemarkParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Remark { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression Replace { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression Append { get; set; }

      [ExpectedExpressionType(typeof(bool))]
      public TypedParsedExpression NewLine { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt muss vom Typ \"{1}\" oder \"{2}\" sein.")]
      TargetTypeError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt \"{1}\" kann nicht bearbeitet werden.")]
      TargetCannotEditError,

      [De ("Befehlsaktivität \"{0}\": Die Beschreibung darf nicht mehr als \"{1}\" Zeichen beinhalten.")]
      ValueTooLong
    }

    public SetRemarkActivityCommandModule ()
        : base ("SetRemark:ActivityCommandClassificationType")
    {
    }

    private string EditRemark (string oldRemark, CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<RemarkParameters> (commandActivity);
      var remarkExpression = parameters.Remark;
      var replaceExpression = parameters.Replace;
      var appendExpression = parameters.Append;
      var newLineExpression = parameters.NewLine;

      var remark = remarkExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var replace = replaceExpression?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      var append = appendExpression?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      var newLine = newLineExpression?.Invoke<bool>(commandActivity, commandActivity.WorkItem);

      string newRemark;

      if (replace == true || string.IsNullOrEmpty (oldRemark))
      {
        newRemark = remark;
      }
      else
      {
        if (append == true)
        {
          if (newLine == true)
            remark = remark.Prepend ("\r\n");
          
          newRemark = oldRemark.Append (remark);
        }
        else
        {
          if (newLine == true)
            remark = remark.Append("\r\n");
          
          newRemark = oldRemark.Prepend (remark);
        }
      }

      ValidateFieldLength (newRemark, 4000, commandActivity);

      return newRemark;
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var target = commandActivity.WorkItem;
      if (target != null && !target.CanEdit (ignoreSuspendable: true))
      {
        throw new ActivityCommandException (
                $"User cannot Edit '{target.ID}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.TargetCannotEditError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    target?.DisplayName));
      }

      switch (commandActivity.WorkItem)
      {
        case FileHandlingContainer fhc:
          fhc.Remark = EditRemark (fhc.Remark, commandActivity);
          break;
        case FileCase fileCase:
          fileCase.WorkInstruction = EditRemark (fileCase.WorkInstruction, commandActivity);
          break;
        default:
          throw new ActivityCommandException (
                  $"WorkItem must be of type '{nameof(FileHandlingContainer)}' or '{nameof(FileCase)}' in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.TargetTypeError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, typeof (FileHandlingContainer).ToLocalizedName(), typeof (FileCase).ToLocalizedName()));
      }

      return true;
    }

    public void ValidateFieldLength (string value, int maxLength, CommandActivity commandActivity)
    {
      if (value?.Length > maxLength)
      {
        throw new ActivityCommandException (
                $"Value can not be longer than {maxLength} in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.ValueTooLong.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, maxLength));
      }
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<RemarkParameters> (commandActivity);
        var replaceExpression = parameters.Replace;
        var appendExpression = parameters.Append;

        var replace = replaceExpression?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
        var append = appendExpression?.Invoke<bool> (commandActivity, commandActivity.WorkItem);

        // Validation during template creation
        if (commandActivity.WorkItem == null && !(commandActivity.EffectiveWorkItemType.IsSubclassOf (typeof (FileHandlingContainer)))
                                             && (commandActivity.EffectiveWorkItemType != typeof (FileCase)))
        {
          throw new ActivityCommandException (
                  $"WorkItemType must be of type '{nameof(FileHandlingContainer)}' or '{nameof(FileCase)}' in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.TargetTypeError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, typeof (FileHandlingContainer).ToLocalizedName(), typeof (FileCase).ToLocalizedName()));
        }

        // Validation in active process
        if (commandActivity.WorkItem != null && !(commandActivity.WorkItem is FileHandlingContainer) && !(commandActivity.WorkItem is FileCase))
        {
          throw new ActivityCommandException (
                  $"WorkItem must be of type '{nameof(FileHandlingContainer)}' or '{nameof(FileCase)}' in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.TargetTypeError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, typeof (FileHandlingContainer).ToLocalizedName(), typeof (FileCase).ToLocalizedName()));
        }

        if (replace == null)
        {
          throw new ActivityCommandException (
                  $"Parameter 'replace' must not be null in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, "replace"));
        }

        if (replace == false && append == null)
        {
          throw new ActivityCommandException (
                  $"Parameter 'append' must not be null in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, "append"));
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