// Name: Aktivität setzen (SetActivity)
// Version: 1.2.3
// Datum: 19.05.2025
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.Linq;
using ActaNova.Domain.Workflow;
using JetBrains.Annotations;
using Remotion.Globalization;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Multilingual.Shared;
using Rubicon.Multilingual;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using Rubicon.Workflow.SecurityManager.Domain;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Multilingual.Extensions;
using System.Globalization;
using System.Reflection;
using Rubicon.Gever.Bund.Domain.Extensions;

namespace Rubicon.Gever.Bund.Domain.Workflow
{
  [VersionInfo ("1.2.3", "19.05.2025", "RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu")]
  [UsedImplicitly]
  public class SetActivityActivityCommandModule : GeverActivityCommandModule
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
    public new enum Localization
    {
      [De ("Befehlsaktivität \"{0}\": Die Nachfolgeaktivität ist im Prozess nicht eindeutig oder nicht vorhanden. Passen Sie den Prozess an.")]
      NextActivityError,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt \"{1}\" kann nicht bearbeitet werden.")]
      TargetCannotEditError,

      [De ("Befehlsaktivität \"{0}\": Der Wert darf nicht mehr als \"{1}\" Zeichen beinhalten.")]
      ValueTooLong,

      [De ("Befehlsaktivität \"{0}\": Befehlsaktivität konnte nicht korrekt ausgeführt werden. Wenden Sie sich an ihren Prozessadministrator.")]
      UnknownError,

      [De (
          "Befehlsaktivität \"{0}\": Der E-Mail-Benachrichtigungstyp \"{1}\" ist ungültig. Folgende Werte stehen zu Verfügung: None, Standard, Important. Gross- und Kleinschreibung beachten.")]
      NotificationEmailTypeInvalidError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" ist schreibgeschützt und darf nicht editiert werden.")]
      PropertyIsReadOnlyError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" muss nach der Auswertung entweder einen Wert enthalten oder entfernt werden.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Version: \"{1}\" - Datum: \"{2}\" - Autor \"{3}\"")]
      ShowVersion
    }

    public class SetActivityParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (TenantUser))]
      [EitherOr (new[] { nameof(User), nameof(Group) })]
      public TypedParsedExpression User { get; set; }

      [ExpectedExpressionType (typeof (TenantGroup))]
      [EitherOr (new[] { nameof(User), nameof(Group) })]
      public TypedParsedExpression Group { get; set; }

      [ExpectedExpressionType (typeof (TenantPosition))]
      public TypedParsedExpression Position { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameDE { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameEN { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameFR { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameIT { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NotificationEmail { get; set; }

      [ExpectedExpressionType (typeof (int?))]
      public TypedParsedExpression SuspendForSeconds { get; set; }

      [ExpectedExpressionType (typeof (double?))]
      public TypedParsedExpression RequiredDurationDays { get; set; }

      [ExpectedExpressionType (typeof (int?))]
      public TypedParsedExpression RequiredDurationMarginPercent { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression AttentionRemark { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression ShowVersion { get; set; }
    }

    public SetActivityActivityCommandModule ()
        : base ("SetActivity:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<SetActivityParameters> (commandActivity);
      var nextActivity = GetNextActivity (commandActivity);

      if (!nextActivity.CanEdit (ignoreSuspendable: true))
      {
        throw new ActivityCommandException (
                $"Activity '{nextActivity.ID}' cannot be edited in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                Localization.TargetCannotEditError.ToLocalizedName().FormatWith (commandActivity.DisplayName, nextActivity?.DisplayName));
      }

      if (parameters.Group != null || parameters.Position != null || parameters.User != null)
        SetRecipient (commandActivity, nextActivity, parameters);

      if (parameters.NameDE != null || parameters.NameEN != null || parameters.NameFR != null || parameters.NameIT != null)
        SetMultilingualName (commandActivity, nextActivity, parameters);

      if (parameters.NotificationEmail != null)
        SetNotificationEmail (commandActivity, nextActivity, parameters);

      if (parameters.SuspendForSeconds != null)
        SetSuspendForSeconds (commandActivity, nextActivity, parameters);

      if (parameters.RequiredDurationDays != null)
        SetRequiredDurationDays (commandActivity, nextActivity, parameters);

      if (parameters.RequiredDurationMarginPercent != null)
        SetRequiredDurationMarginPercent (commandActivity, nextActivity, parameters);

      if (parameters.AttentionRemark != null)
        SetAttentionRemark (commandActivity, nextActivity, parameters);

      nextActivity.CalculateActiveDeadline (false);

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<SetActivityParameters> (commandActivity);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);
        ShowVersion (commandActivity, parameters);
        GetNextActivity (commandActivity);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    private void ShowVersion (CommandActivity commandActivity, SetActivityParameters parameters)
    {
      var showVersion = parameters.ShowVersion?.Invoke<bool> (commandActivity, commandActivity.WorkItem);
      if (showVersion == true)
      {
        var attr = GetType().GetCustomAttribute<VersionInfoAttribute>();
        if (attr != null)
        {
          throw new ActivityCommandException (
                  $"CommandActivity \"{commandActivity.DisplayName}\": Version: \"{attr.Version}\" - Date: \"{attr.Date}\" - Author \"{attr.Author}\"")
              .WithUserMessage (Localization
                  .ShowVersion.ToLocalizedName().FormatWith (
                      commandActivity.DisplayName,
                      attr.Version,
                      attr.Date,
                      attr.Author));
        }
      }
    }

    protected virtual Activity GetNextActivity (CommandActivity commandActivity)
    {
      var nextActivities = commandActivity.GetNextActivitiesAfter (new ActivityBase.GetNextActivitiesSpecification<ActivityBase>());
      if (nextActivities.Count != 1 || !(nextActivities[0] is Activity))
        throw new ActivityCommandException (
                $"Activity after '{commandActivity.ID}' with ID '{nextActivities.FirstOrDefault()?.ID}' is not of type Activity.")
            .WithUserMessage (Localization.NextActivityError.ToLocalizedName().FormatWith (commandActivity.DisplayName));

      return (Activity)nextActivities[0];
    }

    public string CleanInputParameter (CommandActivity commandActivity, TypedParsedExpression nameExpression)
    {
      var name = nameExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      if (name == string.Empty)
        return name;
      return string.IsNullOrWhiteSpace (name) ? null : name;
    }

    public void ValidateMultilingual (CommandActivity commandActivity, MultilingualValueSet valueSet)
    {
      var standardCulture = MultilingualConfigurationSection.Current.StandardDataCultureInfo;
      var isStandardCultureValueSet = !string.IsNullOrEmpty (valueSet.GetValue (standardCulture));
      var hasMultipleAdditionalCultureValuesSet = valueSet.AllValues.Count (v => !v.Culture.Equals (standardCulture) && !string.IsNullOrEmpty (v.Value)) > 1;

      if (!isStandardCultureValueSet && hasMultipleAdditionalCultureValuesSet)
      {
        throw new ActivityCommandException (
                $"If for more than one additional data culture a value is set, then the value for the default data culture has to be set as well, in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                MultilingualValueSet.Localizations.StandardCultureIsRequiredIfMoreThanOneAdditionalCultureIsSet
                    .ToLocalizedName().FormatWith (standardCulture.NativeName));
      }
    }

    public void ValidateFieldLength (CommandActivity commandActivity, string value, int maxLength)
    {
      if (value?.Length > maxLength)
      {
        throw new ActivityCommandException (
                $"Value can not be longer than {maxLength} in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.ValueTooLong.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, maxLength));
      }
    }

    public void SetRecipient (CommandActivity commandActivity, Activity nextActivity, SetActivityParameters parameters)
    {
      var userExpression = parameters.User;
      var groupExpression = parameters.Group;
      var positionExpression = parameters.Position;

      var newRecipientUser = userExpression?.Invoke<TenantUser> (commandActivity, commandActivity.WorkItem);
      var newRecipientGroup = groupExpression?.Invoke<TenantGroup> (commandActivity, commandActivity.WorkItem);
      var newRecipientPosition = positionExpression?.Invoke<TenantPosition> (commandActivity, commandActivity.WorkItem);

      if (!(newRecipientUser == null && newRecipientGroup == null))
      {
        var secRecipient = (SecurityManagerRecipient)nextActivity.Recipient;
        secRecipient.DynamicRecipient = null;
        secRecipient.RecipientSet = null;

        if (newRecipientUser != null)
        {
          secRecipient.User = newRecipientUser;
          secRecipient.Position = null;
          secRecipient.Group = null;
        }
        else
        {
          secRecipient.User = null;
          secRecipient.Position = newRecipientPosition;
          secRecipient.Group = newRecipientGroup;
        }

        nextActivity.Recipient = secRecipient;
      }
    }

    public void SetMultilingualName (CommandActivity commandActivity, Activity nextActivity, SetActivityParameters parameters)
    {
      AssertIsNotReadOnly (commandActivity, nextActivity, nameof(Activity.Name));

      var nameDe = CleanInputParameter (commandActivity, parameters.NameDE);
      var nameEn = CleanInputParameter (commandActivity, parameters.NameEN);
      var nameFr = CleanInputParameter (commandActivity, parameters.NameFR);
      var nameIt = CleanInputParameter (commandActivity, parameters.NameIT);

      ValidateFieldLength (commandActivity, nameDe, 400);
      ValidateFieldLength (commandActivity, nameFr, 400);
      ValidateFieldLength (commandActivity, nameIt, 400);
      ValidateFieldLength (commandActivity, nameEn, 400);

      if (nameDe != null)
        nextActivity.SetMultilingualValue (f => f.Name, nameDe, CultureInfo.GetCultureInfo ("de"));

      if (nameEn != null)
        nextActivity.SetMultilingualValue (f => f.Name, nameEn, CultureInfo.GetCultureInfo ("en"));

      if (nameFr != null)
        nextActivity.SetMultilingualValue (f => f.Name, nameFr, CultureInfo.GetCultureInfo ("fr"));

      if (nameIt != null)
        nextActivity.SetMultilingualValue (f => f.Name, nameIt, CultureInfo.GetCultureInfo ("it"));

      var valueSet = nextActivity.GetMultilingualValueSet (f => f.Name);
      ValidateMultilingual (commandActivity, valueSet);
    }

    public void SetNotificationEmail (CommandActivity commandActivity, Activity nextActivity, SetActivityParameters parameters)
    {
      AssertIsNotReadOnly (commandActivity, nextActivity, nameof(Activity.NotificationEmail));
      string notificationEmailType;
      try
      {
        notificationEmailType = parameters.NotificationEmail?.Invoke<string> (commandActivity, commandActivity.WorkItem).Trim();
      }
      catch (Exception ex) {
        throw new ActivityCommandException(
            $"Parameter {nameof(parameters.NotificationEmail)} must not be empty after invocation or must be removed in activity '{commandActivity.ID}'").WithUserMessage(
            Localization
                .ParameterNullError.ToLocalizedName().FormatWith(commandActivity.DisplayName, nameof(parameters.NotificationEmail)));
      }
      
      var notificationEmailTypeValid = Enum.TryParse (notificationEmailType, out NotificationEmailType emailType);
      if (!notificationEmailTypeValid)
      {
        throw new ActivityCommandException (
                $"Notification email type '{notificationEmailType}' is invalid in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.NotificationEmailTypeInvalidError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, notificationEmailType));
      }

      nextActivity.NotificationEmail = emailType;
    }

    public void SetSuspendForSeconds (CommandActivity commandActivity, Activity nextActivity, SetActivityParameters parameters)
    {
      AssertIsNotReadOnly (commandActivity, nextActivity, nameof(Activity.SuspendForSeconds));
      var seconds = parameters.SuspendForSeconds?.Invoke<int?> (commandActivity, commandActivity.WorkItem);
      if (seconds == null)
      {
        throw new ActivityCommandException(
            $"Parameter {nameof(parameters.SuspendForSeconds)} must not be empty or must be removed in activity '{commandActivity.ID}'").WithUserMessage(
            Localization
                .ParameterNullError.ToLocalizedName().FormatWith(commandActivity.DisplayName, nameof(parameters.SuspendForSeconds)));
      }

      var validRange = new SuspendForSecondsRange { Seconds = (double)seconds };
      new ParameterValidator (commandActivity).ValidateParameters (validRange);
      nextActivity.SuspendForSeconds = seconds;
    }

    public void SetRequiredDurationDays (CommandActivity commandActivity, Activity nextActivity, SetActivityParameters parameters)
    {
      AssertIsNotReadOnly (commandActivity, nextActivity, nameof(Activity.RequiredDurationDays));
      var requiredDurationDays = parameters.RequiredDurationDays?.Invoke<double?> (commandActivity, commandActivity.WorkItem);
      if (requiredDurationDays == null)
      {
        throw new ActivityCommandException(
            $"Parameter {nameof(parameters.RequiredDurationDays)} must not be empty or must be removed in activity '{commandActivity.ID}'").WithUserMessage(
            Localization
                .ParameterNullError.ToLocalizedName().FormatWith(commandActivity.DisplayName, nameof(parameters.RequiredDurationDays)));
      }

      var validRange = new RequiredDurationDaysRange { Days = (double)requiredDurationDays };
      new ParameterValidator (commandActivity).ValidateParameters (validRange);
      nextActivity.RequiredDurationDays = (double)requiredDurationDays;
    }

    public void SetRequiredDurationMarginPercent (CommandActivity commandActivity, Activity nextActivity, SetActivityParameters parameters)
    {
      AssertIsNotReadOnly (commandActivity, nextActivity, nameof(Activity.RequiredDurationMarginPercent));
      var requiredDurationMarginPercent = parameters.RequiredDurationMarginPercent?.Invoke<int?>(commandActivity, commandActivity.WorkItem);
      if (requiredDurationMarginPercent == null)
      {
        throw new ActivityCommandException(
                $"Parameter {nameof(parameters.RequiredDurationMarginPercent)} must not be empty or must be removed in activity '{commandActivity.ID}'")
            .WithUserMessage(Localization
                .ParameterNullError.ToLocalizedName().FormatWith(commandActivity.DisplayName, nameof(parameters.RequiredDurationMarginPercent)));
      }

      var validRange = new RequiredDurationMarginPercentRange { MarginPercent = (double)requiredDurationMarginPercent };
      new ParameterValidator (commandActivity).ValidateParameters (validRange);
      nextActivity.RequiredDurationMarginPercent = (int)requiredDurationMarginPercent;
    }

    public void SetAttentionRemark (CommandActivity commandActivity, Activity nextActivity, SetActivityParameters parameters)
    {
      AssertIsNotReadOnly (commandActivity, nextActivity, nameof(Activity.AttentionRemark));
      var attentionRemark = parameters.AttentionRemark.Invoke<string> (commandActivity, commandActivity.WorkItem);
      nextActivity.AttentionRemark = attentionRemark;
    }

    public void AssertIsNotReadOnly (CommandActivity commandActivity, Activity activity, string propertyIdentifier)
    {
      if (activity.GetPropertyState (propertyIdentifier).IsReadOnly == true)
      {
        throw new ActivityCommandException (
                $"The property '{propertyIdentifier}' is ReadOnly in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.PropertyIsReadOnlyError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, propertyIdentifier));
      }
    }

    private class SuspendForSecondsRange
    {
      [Range (0, int.MaxValue)]
      public double Seconds { get; set; }
    }

    private class RequiredDurationDaysRange
    {
      [Range (0, 36500)]
      public double Days { get; set; }
    }

    private class RequiredDurationMarginPercentRange
    {
      [Range (0, 100)]
      public double MarginPercent { get; set; }
    }
  }
}