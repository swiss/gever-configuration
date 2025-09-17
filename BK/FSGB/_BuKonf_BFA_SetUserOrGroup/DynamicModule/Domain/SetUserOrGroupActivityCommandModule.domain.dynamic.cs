// Name: Benutzer oder Gruppe setzen (SetUserOrGroup)
// Version: 2.0.2
// Datum: 03.11.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu

using System;
using System.Diagnostics.CodeAnalysis;
using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using JetBrains.Annotations;
using Remotion.Globalization;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Mixins;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using Rubicon.Workflow.SecurityManager.Domain;

namespace Rubicon.Gever.Bund.Domain.Workflow
{
  [UsedImplicitly]
  [ExcludeFromCodeCoverage]
  public class SetUserOrGroupActivityCommandModule : GeverActivityCommandModule
  {
    [LocalizationEnum]
    public enum Localization
    {
      [De ("Befehlsaktivität \"{0}\": Beim Auswerten des Ausdrucks des Parameters \"{1}\" konnte kein Benutzer oder keine Gruppe ermittelt werden.")]
      EmptyExpressionResultError,

      [De ("Befehlsaktivität \"{0}\": Die Befehlsaktivität konnte nicht korrekt ausgeführt werden. Wenden Sie sich an Ihren Prozessadministrator.")]
      UnknownError,

      [De ("Befehlsaktivität \"{0}\": Der Ausdruck im Parameter \"Target\" liefert kein Ergebnis. Wenden Sie sich an Ihren Prozessadministrator.")]
      TargetExpressionEmptyResultError,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt muss vom Typ \"{1}\" sein.")]
      TargetTypeError,

      [De ("Befehlsaktivität \"{0}\": Mindestens \"{1}\", \"{2}\", \"{3}\", \"{4}\", \"{5}\" oder \"{6}\" muss einen Wert haben.")]
      RequiredParameterError,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt \"{1}\" kann nicht bearbeitet werden.")]
      TargetCannotEditError,

      [De ("Befehlsaktivität \"{0}\": SignerLeft und SignerRight sind nur auf Dokumenten verfügbar.")]
      WrongTargetForSigner
    }

    public class UserOrGroupParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (FileHandlingObject))]
      public TypedParsedExpression Target { get; set; }

      [ExpectedExpressionType (typeof (TenantUser))]
      public TypedParsedExpression LeadingUserExpression { get; set; }

      [ExpectedExpressionType (typeof (TenantGroup))]
      public TypedParsedExpression LeadingGroupExpression { get; set; }

      [ExpectedExpressionType (typeof (TenantUser))]
      public TypedParsedExpression ProcessOwnerExpression { get; set; }

      [ExpectedExpressionType (typeof (TenantUser))]
      public TypedParsedExpression SignerLeftExpression { get; set; }

      [ExpectedExpressionType (typeof (TenantUser))]
      public TypedParsedExpression SignerRightExpression { get; set; }

      [ExpectedExpressionType (typeof (TenantUser))]
      public TypedParsedExpression CreatorExpression { get; set; }
    }

    public SetUserOrGroupActivityCommandModule ()
        : base ("SetUserOrGroup:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<UserOrGroupParameters> (commandActivity);
      var leadingUserExpression = parameters.LeadingUserExpression;
      var leadingGroupExpression = parameters.LeadingGroupExpression;
      var processOwnerExpression = parameters.ProcessOwnerExpression;
      var signerLeftExpression = parameters.SignerLeftExpression;
      var signerRightExpression = parameters.SignerRightExpression;
      var creatorExpression = parameters.CreatorExpression;

      var leadingUser = leadingUserExpression?.Invoke<TenantUser> (commandActivity, commandActivity.WorkItem);
      var leadingGroup = leadingGroupExpression?.Invoke<TenantGroup> (commandActivity, commandActivity.WorkItem);
      var targetFho = GetTargetFromWorkItemOrParameter (commandActivity, parameters.Target);
      var processOwner = processOwnerExpression?.Invoke<TenantUser> (commandActivity, commandActivity.WorkItem);
      var signerLeft = signerLeftExpression?.Invoke<TenantUser> (commandActivity, commandActivity.WorkItem);
      var signerRight = signerRightExpression?.Invoke<TenantUser> (commandActivity, commandActivity.WorkItem);
      var creator = creatorExpression?.Invoke<TenantUser> (commandActivity, commandActivity.WorkItem);

      if (processOwnerExpression != null && processOwner == null)
      {
        throw new ActivityCommandException ($"Expression returned null instead of TenantUser in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.EmptyExpressionResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.ProcessOwnerExpression)));
      }

      if (leadingUserExpression != null && leadingUser == null)
      {
        throw new ActivityCommandException ($"Expression returned null instead of TenantUser in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.EmptyExpressionResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.LeadingUserExpression)));
      }

      if (leadingGroupExpression != null && leadingGroup == null)
      {
        throw new ActivityCommandException ($"Expression returned null instead of TenantGroup in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.EmptyExpressionResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.LeadingGroupExpression)));
      }

      if (signerLeftExpression != null && signerLeft == null)
      {
        throw new ActivityCommandException ($"Expression returned null instead of TenantGroup in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.EmptyExpressionResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.SignerLeftExpression)));
      }

      if (signerRightExpression != null && signerRight == null)
      {
        throw new ActivityCommandException ($"Expression returned null instead of TenantGroup in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.EmptyExpressionResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.SignerRightExpression)));
      }

      if (creatorExpression != null && creator == null)
      {
        throw new ActivityCommandException ($"Expression returned null instead of TenantGroup in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.EmptyExpressionResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.CreatorExpression)));
      }

      if ((signerLeft != null || signerRight != null) && !(targetFho is IGeverSigner))
      {
        throw new ActivityCommandException ($"Expression returned null instead of TenantGroup in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.WrongTargetForSigner.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName));
      }

      if (leadingUser != null)
        targetFho.LeadingUser = leadingUser;
      
      if (leadingGroup != null)
        targetFho.LeadingGroup = leadingGroup;

      if (processOwner != null)
      {
        var process = commandActivity.GetParentActivity().RootActivity;

        if (process is ProcessActivity processActivity)
        {
          ((IProcessActivityMixin)processActivity).ProcessOwner = processOwner;
        }
      }

      if (signerLeft != null)
        ((IGeverSigner)targetFho).SignerLeft = signerLeft;

      if (signerRight != null)
        ((IGeverSigner)targetFho).SignerRight = signerRight;

      if (creator != null)
        targetFho.CreatedBy = creator;
      
      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var parameters = ActivityCommandParser.Parse<UserOrGroupParameters> (commandActivity);
        if (parameters.LeadingUserExpression == null && parameters.LeadingGroupExpression == null && parameters.ProcessOwnerExpression == null &&
            parameters.SignerLeftExpression == null && parameters.SignerRightExpression == null && parameters.CreatorExpression == null)
        {
          throw new ActivityCommandException ($"Error while parsing activity '{commandActivity.ID}'.")
              .WithUserMessage (Localization.RequiredParameterError.ToLocalizedName().FormatWith (commandActivity.DisplayName,
                  nameof(parameters.LeadingUserExpression),
                  nameof(parameters.LeadingGroupExpression),
                  nameof(parameters.ProcessOwnerExpression),
                  nameof(parameters.SignerLeftExpression),
                  nameof(parameters.SignerRightExpression),
                  nameof(parameters.CreatorExpression)
                  ));
        }
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    protected FileHandlingObject GetTargetFromWorkItemOrParameter (CommandActivity commandActivity, TypedParsedExpression targetExpression)
    {
      FileHandlingObject target = null;

      if (targetExpression != null)
      {
        target = targetExpression.Invoke<FileHandlingObject> (commandActivity, commandActivity.WorkItem);

        if (target == null)
        {
          throw new ActivityCommandException ($"TargetExpression returned null in activity '{commandActivity.ID}'.")
              .WithUserMessage (Localization.TargetExpressionEmptyResultError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }
      }
      else if (commandActivity.WorkItem != null)
      {
        target = commandActivity.WorkItem as FileHandlingObject;

        if (target == null)
        {
          throw new ActivityCommandException (
                  $"Das Zielobjekt muss vom Typ '{nameof(FileHandlingObject)}' sein in activity '{commandActivity.ID}'.")
              .WithUserMessage (Localization.TargetTypeError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, typeof (FileHandlingObject).ToLocalizedName()));
        }
      }

      if (target != null && !target.CanEdit (ignoreSuspendable: true))
      {
        throw new ActivityCommandException (
                $"User cannot Edit '{target.ID}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                Localization.TargetCannotEditError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    target?.DisplayName));
      }

      return target;
    }
  }
}