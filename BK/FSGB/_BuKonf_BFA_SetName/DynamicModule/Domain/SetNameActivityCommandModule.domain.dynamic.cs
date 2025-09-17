// Name: Objektname setzen (setName)
// Version: 2.0.0
// Datum: 04.07.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu

using System;
using System.Globalization;
using System.Linq;
using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Multilingual;
using Rubicon.Multilingual.Extensions;
using Rubicon.Multilingual.Shared;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  public class SetNameActivityCommandModule : GeverActivityCommandModule
  {
    public class NameParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Name { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameDE { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameEN { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameFR { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression NameIT { get; set; }

      [ExpectedExpressionType (typeof (BaseActaNovaObject))]
      public TypedParsedExpression Target { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt darf nicht vom Typ \"{1}\" sein.")]
      TargetTypeError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Mindestens einer der Parameter \"{1}\", \"{2}\", \"{3}\", \"{4}\", oder \"{5}\" darf nicht null sein.")]
      MultilingualParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt \"{1}\" kann nicht bearbeitet werden.")]
      TargetCannotEditError,

      [De ("Befehlsaktivität \"{0}\": Der Wert darf nicht mehr als \"{1}\" Zeichen beinhalten.")] 
      ValueTooLong
    }

    public SetNameActivityCommandModule ()
        : base ("SetName:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<NameParameters> (commandActivity);
      var nameExpression = parameters.Name;
      var nameDeExpression = parameters.NameDE;
      var nameEnExpression = parameters.NameEN;
      var nameFrExpression = parameters.NameFR;
      var nameItExpression = parameters.NameIT;
      var targetExpression = parameters.Target;

      var name = nameExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var nameDe = nameDeExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var nameEn = nameEnExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var nameFr = nameFrExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var nameIt = nameItExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var target = targetExpression?.Invoke<BaseActaNovaObject> (commandActivity, commandActivity.WorkItem);

      if (target == null)
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(NameParameters.Target)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.ParameterNullError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, nameof(NameParameters.Target)));
      }

      if (!target.CanEdit (ignoreSuspendable: true))
      {
        throw new ActivityCommandException (
                $"User cannot Edit '{target.ID}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.TargetCannotEditError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    target?.DisplayName));
      }

      var targetType = target.GetPublicDomainObjectType();

      if (name == null && !targetType.IsSubclassOf (typeof (BaseFile)) && targetType != typeof (FileCase))
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(NameParameters.Name)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.ParameterNullError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, nameof(NameParameters.Name)));
      }

      if (name == null
          && nameDe == null
          && nameEn == null
          && nameFr == null
          && nameIt == null
          && (targetType.IsSubclassOf (typeof (BaseFile)) || targetType == typeof (FileCase)))
      {
        throw new ActivityCommandException (
                $"At least one of the parameters '{nameof(NameParameters.Name)}', {nameof(NameParameters.NameDE)}', '{nameof(NameParameters.NameEN)}', '{nameof(NameParameters.NameFR)}' or '{nameof(NameParameters.NameIT)}' must not be null in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.MultilingualParameterNullError.ToLocalizedName()
                    .FormatWith (
                        commandActivity.DisplayName,
                        nameof(NameParameters.Name),
                        nameof(NameParameters.NameDE),
                        nameof(NameParameters.NameFR),
                        nameof(NameParameters.NameIT),
                        nameof(NameParameters.NameEN)));
      }

      ValidateFieldLength (name, 400, commandActivity);
      ValidateFieldLength (nameDe, 400, commandActivity);
      ValidateFieldLength (nameFr, 400, commandActivity);
      ValidateFieldLength (nameIt, 400, commandActivity);
      ValidateFieldLength (nameEn, 400, commandActivity);

      if (target is BaseFile file)
      {
        if (nameDe != null)
          file.SetMultilingualValue (f => f.Title, nameDe, CultureInfo.GetCultureInfo ("de"));
        else if (name != null)
          file.SetMultilingualValue (f => f.Title, name, CultureInfo.GetCultureInfo ("de"));

        if (nameEn != null)
          file.SetMultilingualValue (f => f.Title, nameEn, CultureInfo.GetCultureInfo ("en"));

        if (nameFr != null)
          file.SetMultilingualValue (f => f.Title, nameFr, CultureInfo.GetCultureInfo ("fr"));

        if (nameIt != null)
          file.SetMultilingualValue (f => f.Title, nameIt, CultureInfo.GetCultureInfo ("it"));

        var valueSet = file.GetMultilingualValueSet (f => f.Title);
        ValidateMultilingual (valueSet, commandActivity);
      }
      else if (target is FileCase fileCase)
      {
        if (nameDe != null)
          fileCase.SetMultilingualValue (f => f.Title, nameDe, CultureInfo.GetCultureInfo ("de"));
        else if (name != null)
          fileCase.SetMultilingualValue (f => f.Title, name, CultureInfo.GetCultureInfo ("de"));

        if (nameEn != null)
          fileCase.SetMultilingualValue (f => f.Title, nameEn, CultureInfo.GetCultureInfo ("en"));

        if (nameFr != null)
          fileCase.SetMultilingualValue (f => f.Title, nameFr, CultureInfo.GetCultureInfo ("fr"));

        if (nameIt != null)
          fileCase.SetMultilingualValue (f => f.Title, nameIt, CultureInfo.GetCultureInfo ("it"));

        var valueSet = fileCase.GetMultilingualValueSet (f => f.Title);
        ValidateMultilingual (valueSet, commandActivity);
      }
      else if (target is BaseIncoming incoming)
      {
        incoming.Subject = name;
      }
      else if (target is Settlement settlement)
      {
        settlement.ActiveContent.Name = name;
      }
      else if (target is Document document)
      {
        document.ActiveContent.Name = name;
      }
      else if (target is SpecialdataHostingBaseDataObject dataObject)
      {
        dataObject.Name = name;
      }
      else
      {
        throw new ActivityCommandException (
                $"Target type '{targetType}' is not supported in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.TargetTypeError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, targetType));
      }

      return true;
    }

    public void ValidateMultilingual (MultilingualValueSet valueSet, CommandActivity commandActivity)
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
        var parameters = ActivityCommandParser.Parse<NameParameters> (commandActivity);
        var nameExpression = parameters.Name;
        var nameDeExpression = parameters.NameDE;
        var nameEnExpression = parameters.NameEN;
        var nameFrExpression = parameters.NameFR;
        var nameItExpression = parameters.NameIT;
        var targetExpression = parameters.Target;

        if (targetExpression == null)
        {
          throw new ActivityCommandException (
                  $"Parameter '{nameof(NameParameters.Target)}' must not be null in activity '{commandActivity.ID}'.")
              .WithUserMessage (
                  ModuleLocalization.ParameterNullError.ToLocalizedName()
                      .FormatWith (commandActivity.DisplayName, nameof(NameParameters.Target)));
        }

        if (nameExpression == null
            && nameDeExpression == null
            && nameEnExpression == null
            && nameFrExpression == null
            && nameItExpression == null)
        {
          throw new ActivityCommandException (
                  $"At least one of the parameters '{nameof(NameParameters.Name)}', {nameof(NameParameters.NameDE)}', '{nameof(NameParameters.NameEN)}', '{nameof(NameParameters.NameFR)}' or '{nameof(NameParameters.NameIT)}' must not be null in activity '{commandActivity.ID}'.")
              .WithUserMessage (
                  ModuleLocalization.MultilingualParameterNullError.ToLocalizedName()
                      .FormatWith (
                          commandActivity.DisplayName,
                          nameof(NameParameters.Name),
                          nameof(NameParameters.NameDE),
                          nameof(NameParameters.NameFR),
                          nameof(NameParameters.NameIT),
                          nameof(NameParameters.NameEN)));
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