// Name: Prozesshinweis eintragen (AddActivityRemark)
// Version: 1.0.0
// Datum: 19.05.2025
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.Linq;
using System.Reflection;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  [VersionInfo ("1.0.0", "19.05.2025", "RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu")]
  public class AddActivityRemarkActivityCommandModule : GeverActivityCommandModule
  {
    [AttributeUsage (AttributeTargets.Class, Inherited = false)]
    private class VersionInfoAttribute : Attribute
    {
      public string Version { get; }
      public string Date { get; }
      public string Author { get; }

      public VersionInfoAttribute (string version, string date, string author)
      {
        Version = version;
        Date = date;
        Author = author;
      }
    }

    [LocalizationEnum]
    private enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Die Nachfolgeaktivität ist im Prozess nicht eindeutig oder nicht vorhanden. Passen Sie den Prozess an.")]
      NextActivityError,

      [De ("Befehlsaktivität \"{0}\": Beim Setzen des Arbeits- oder Erledigungshinweises ist ein Fehler aufgetreten.")]
      SettingActivityRemarkError,

      [De ("Befehlsaktivität \"{0}\": Version: \"{1}\" - Datum: \"{2}\" - Autor \"{3}\"")]
      ShowVersion
    }

    private class Parameters : GeverActivityCommandParameters
    {
      [RequiredOneOfMany (new[] { nameof(AttentionRemark), nameof(CompletionNote) })]
      [EitherOr (new[] { nameof(AttentionRemark), nameof(CompletionNote) })]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression AttentionRemark { get; set; }

      [RequiredOneOfMany (new[] { nameof(AttentionRemark), nameof(CompletionNote) })]
      [EitherOr (new[] { nameof(AttentionRemark), nameof(CompletionNote) })]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression CompletionNote { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression ShowVersion { get; set; }
    }

    public AddActivityRemarkActivityCommandModule ()
        : base ("AddActivityRemark:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<Parameters> (commandActivity);
      var nextActivity = GetNextActivity (commandActivity);
      try
      {
        String message;
        if (parameters.AttentionRemark != null)
        {
          message = parameters.AttentionRemark?.Invoke<string> (commandActivity, commandActivity.WorkItem);
          nextActivity.AttentionRemark = message;
        }
        else
        {
          message = parameters.CompletionNote?.Invoke<string> (commandActivity, commandActivity.WorkItem);
          ((IActaNovaActivity)nextActivity).CompletionNote = message;
        }
      }
      catch (Exception)
      {
        throw new ActivityCommandException ($"Setting the AttentionRemark or the CompletionNote failed in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.SettingActivityRemarkError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }

      return true;
    }

    protected virtual Activity GetNextActivity (CommandActivity commandActivity)
    {
      var nextActivities = commandActivity.GetNextActivitiesAfter (new ActivityBase.GetNextActivitiesSpecification<ActivityBase>());
      if (nextActivities.Count != 1 || !(nextActivities[0] is Activity))
        throw new ActivityCommandException (
                $"Activity after '{commandActivity.ID}' with ID '{nextActivities.FirstOrDefault()?.ID}' is not of type Activity, in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.NextActivityError.ToLocalizedName().FormatWith (commandActivity.DisplayName));

      return (Activity)nextActivities[0];
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<Parameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);
        ShowVersion (commandActivity, parameters);

        return null;
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }
    }

    private void ShowVersion (CommandActivity commandActivity, Parameters parameters)
    {
      var showVersion = parameters.ShowVersion?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      if (showVersion == true)
      {
        var attr = GetType().GetCustomAttribute<VersionInfoAttribute>();
        if (attr != null)
        {
          throw new ActivityCommandException (
                  $"CommandActivity \"{commandActivity.DisplayName}\": Version: \"{attr.Version}\" - Date: \"{attr.Date}\" - Author \"{attr.Author}\"")
              .WithUserMessage (ModuleLocalization
                  .ShowVersion.ToLocalizedName().FormatWith (
                      commandActivity.DisplayName,
                      attr.Version,
                      attr.Date,
                      attr.Author));
        }
      }
    }
  }
}