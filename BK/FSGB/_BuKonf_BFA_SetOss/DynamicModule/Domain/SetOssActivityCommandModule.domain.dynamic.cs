// Name: Sicherheitseinstellungen setzen (setOss)
// Version: 2.2.2
// Datum: 21.12.2023
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.Collections.Generic;
using System.Linq;
using ActaNova.Domain;
using ActaNova.Domain.AccessDefinition;
using ActaNova.Domain.Utilities;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Remotion.Security;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  public class SetOssActivityCommandModule : GeverActivityCommandModule, ITemplateOnlyActivityCommandModule
  {
    public class SetOssParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (BaseSecuredObjectWithOSS))]
      public TypedParsedExpression Target { get; set; }

      public string SetSecurityAccessDefinitionOrigin { get; set; }

      public string SetSecurityAccessType { get; set; }

      public string SetWorkflowAccessType { get; set; }

      public string SetNotEntitledAccessType { get; set; }

      [ExpectedExpressionType (typeof (TenantBoundOrganizationalStructureObject))]
      public TypedParsedExpression AddUserOrGroup { get; set; }

      [ExpectedExpressionType (typeof (TenantBoundOrganizationalStructureObject))]
      public TypedParsedExpression RemoveUserOrGroup { get; set; }

      public string Access { get; set; }

      public string SetWorkflowAccessUserCurrentlyType { get; set; }

      public string SetWorkflowAccessGroupCurrentlyType { get; set; }

      public string SetWorkflowAccessUserPreviouslyType { get; set; }

      public string SetWorkflowAccessGroupPreviouslyType { get; set; }

      [ExpectedExpressionType (typeof (TenantGroup))]
      public TypedParsedExpression NotificationRecipient { get; set; }

      public bool DisableNotification { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt muss vom Typ \"{1}\" sein.")]
      TargetTypeError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht null sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Der Ausdruck im Parameter \"{1}\" liefert kein Ergebnis.")]
      ExpressionEmptyResultError,

      [De ("Befehlsaktivität \"{0}\": Die Zugriffsregel \"{1}\" konnte nicht gefunden werden.")]
      UnknownSecurityAccessDefinitionType,

      [De ("Befehlsaktivität \"{0}\": Die Sicherheitsherkunft \"{1}\" konnte nicht gefunden werden.")]
      UnknownSecurityAccessDefinitionOrigin,

      [De ("Befehlsaktivität \"{0}\": Das Zielobjekt ist kein direkt untergeordnetes Objekt des WorkItem.")]
      TargetIsNoChildOfWorkItem,

      [De ("Befehlsaktivität \"{0}\": Die Regel für Workflow-Empfänger \"{1}\" konnte nicht gefunden werden.")]
      UnknownWorkflowAccesType,

      [De ("Befehlsaktivität \"{0}\": Die Berechtigungsrolle \"{1}\" konnte nicht gefunden werden.")]
      UnknownSecurityAccessModeType,

      [De ("Befehlsaktivität \"{0}\": Es wurde kein Eintrag für \"{1}\" mit der Berechtigungsrolle \"{2}\" gefunden.")]
      OssEntryNotFound,

      [De ("Befehlsaktivität \"{0}\": Der OSS Eintrag für \"{1}\" mit der Berechtigungsrolle \"{2}\" konnte nicht entfernt werden.")]
      OssEntryNotRemoved,

      [De ("Befehlsaktivität \"{0}\": Nach dem Speichern hat niemand mehr Zugriff auf das Objekt. Ändern Sie die für die Sicherheit relevanten Einstellungen.")]
      NoAccessAfterSaveError,

      [De ("Befehlsaktivität \"{0}\": Es können nicht mehrere Funktionen gleichzeitig verwendet werden. In diesem Fall sind mehrere Aktivitäten vorzusehen.")]
      MultipleFunctionsError,

      [De ("Befehlsaktivität \"{0}\": Die Parameterkombination ist ungültig.")]
      InvalidParameterError,

      [De ("Der Prozess \"{0}\" hat die OSS angepasst!")]
      OssEditedMesssage,

      [De ("Die Aktivität muss schreibgeschützt in der Prozessinstanz sein.")]
      InstanceNotReadOnlyError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht leer sein, wenn die Benachrichtigungen aktiviert sind.")]
      NotificationRecipientNullError,
    }

    public SetOssActivityCommandModule ()
        : base ("SetOss:ActivityCommandClassificationType")
    {
    }

    private BaseSecuredObjectWithOSS _originalTargetObject;

    public override bool Execute (CommandActivity commandActivity)
    {
      var result = base.Execute (commandActivity);

      // At least one user must still have access after OSS changes
      commandActivity.Committing += delegate
      {
        var ossSetting = _originalTargetObject.EffectiveObjectSpecificSecuritySetting;

        if (ossSetting == null)
          return;

        if (ossSetting.SecurityAccessDefinitionType != SecurityAccessDefinitionType.Values.Restricted())
          return;

        if (ossSetting.WorkflowAccessType != WorkflowSecurityAccessDefinitionType.Values.Detailed()
            && ossSetting.WorkflowAccessType != WorkflowSecurityAccessDefinitionType.Values.NoAccess())
          return;

        if (ossSetting.ObjectSpecificSecurityEntries.Any (entry => entry.Access == SecurityAccessModeType.Values.DefineSecurity()))
          return;

        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.").WithUserMessage (
            ModuleLocalization.NoAccessAfterSaveError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      };
      return result;
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      using (SecurityFreeSection.Activate())
      {
        var parameters = ActivityCommandParser.Parse<SetOssParameters> (commandActivity);

        var target = GetTargetFromWorkItemOrParam (commandActivity, parameters);

        _originalTargetObject = target;

        var disableNotification = parameters.DisableNotification;
        if (disableNotification == false)
        {
          var notificationRecipient = parameters.NotificationRecipient.Invoke<TenantGroup> (commandActivity, commandActivity.WorkItem);
          if (notificationRecipient == null)
          {
            throw new ActivityCommandException ($"NotificationRecipient returned null in activity '{commandActivity.ID}'.")
                .WithUserMessage (ModuleLocalization.ExpressionEmptyResultError.ToLocalizedName()
                    .FormatWith (commandActivity.DisplayName, nameof(parameters.NotificationRecipient)));
          }

          SendThroughWorkflowModule.Value (
              target,
              notificationRecipient,
              ModuleLocalization.OssEditedMesssage.ToLocalizedName().FormatWith (((ProcessActivity)commandActivity.RootActivity).DisplayName),
              true,
              null);
        }

        // Set Security Access Definition Origin
        var securityAccessDefinitionOrigin = parameters.SetSecurityAccessDefinitionOrigin;
        if (securityAccessDefinitionOrigin != null)
        {
          SetSecurityAccessDefinitionOrigin (commandActivity, securityAccessDefinitionOrigin, target);
          return true;
        }

        // Set Security Access Type
        var setSecurityAccessType = parameters.SetSecurityAccessType;
        if (setSecurityAccessType != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          SetSecurityAccessType (commandActivity, setSecurityAccessType, target);
          return true;
        }

        // Set Workflow Access Type
        var setWorkflowAccessType = parameters.SetWorkflowAccessType;
        if (setWorkflowAccessType != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          SetWorkflowAccessType (commandActivity, setWorkflowAccessType, target);
          return true;
        }

        // Set Not Entitled
        var setNotEntitledType = parameters.SetNotEntitledAccessType;
        if (setNotEntitledType != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          SetNotEntitledAccessType (commandActivity, setNotEntitledType, target);
          return true;
        }

        // Add User or Group
        var addUserOrGroup = parameters.AddUserOrGroup;
        if (addUserOrGroup != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          AddUserOrGroup (commandActivity, parameters, target);
          return true;
        }

        // Remove User or Group
        var removeUserOrGroup = parameters.RemoveUserOrGroup;
        if (removeUserOrGroup != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          RemoveUserOrGroup (commandActivity, parameters, target);
          return true;
        }

        // Set Workflow Access User Currently Type
        var setWorkflowAccessUserCurrentlyType = parameters.SetWorkflowAccessUserCurrentlyType;
        if (setWorkflowAccessUserCurrentlyType != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          SetWorkflowAccessType (commandActivity, "Detailed", target);
          SetWorkflowAccessUserCurrentlyType (commandActivity, setWorkflowAccessUserCurrentlyType, target);
          return true;
        }

        // Set Workflow Access Group Currently Type
        var setWorkflowAccessGroupCurrentlyType = parameters.SetWorkflowAccessGroupCurrentlyType;
        if (setWorkflowAccessGroupCurrentlyType != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          SetWorkflowAccessType (commandActivity, "Detailed", target);
          SetWorkflowAccessGroupCurrentlyType (commandActivity, setWorkflowAccessGroupCurrentlyType, target);
          return true;
        }

        // Set Workflow Access User Previously Type
        var setWorkflowAccessUserPreviouslyType = parameters.SetWorkflowAccessUserPreviouslyType;
        if (setWorkflowAccessUserPreviouslyType != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          SetWorkflowAccessType (commandActivity, "Detailed", target);
          SetWorkflowAccessUserPreviouslyType (commandActivity, setWorkflowAccessUserPreviouslyType, target);
          return true;
        }

        // Set Workflow Access Group Previously Type
        var setWorkflowAccessGroupPreviouslyType = parameters.SetWorkflowAccessGroupPreviouslyType;
        if (setWorkflowAccessGroupPreviouslyType != null)
        {
          target.SecurityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.Self();
          SetWorkflowAccessType (commandActivity, "Detailed", target);
          SetWorkflowAccessGroupPreviouslyType (commandActivity, setWorkflowAccessGroupPreviouslyType, target);
          return true;
        }

        throw new ActivityCommandException ($"Error while parsing activity'{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.InvalidParameterError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }
    }

    private static void SetNotEntitledAccessType (CommandActivity commandActivity, string setNotEntitledAccessType, BaseSecuredObjectWithOSS target)
    {
      try
      {
        var notEntitledAccessType = SecurityAccessModeType.Values.GetValueInfoByID (setNotEntitledAccessType).Value;
        target.VisibleNotEntitled = notEntitledAccessType;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessModeType.ToLocalizedName().FormatWith (commandActivity.DisplayName, setNotEntitledAccessType));
      }
    }

    private static void SetWorkflowAccessType (CommandActivity commandActivity, string setWorkflowAccessType, BaseSecuredObjectWithOSS target)
    {
      try
      {
        var workflowAccessType = WorkflowSecurityAccessDefinitionType.Values.GetValueInfoByID (setWorkflowAccessType).Value;
        target.VisibleWorkflowAccessType = workflowAccessType;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownWorkflowAccesType.ToLocalizedName().FormatWith (commandActivity.DisplayName, setWorkflowAccessType));
      }
    }

    private static void SetSecurityAccessType (CommandActivity commandActivity, string setSecurityAccessType, BaseSecuredObjectWithOSS target)
    {
      try
      {
        var securityAccessDefinitionType = SecurityAccessDefinitionType.Values.GetValueInfoByID (setSecurityAccessType).Value;
        target.VisibleSecurityAccessDefinitionType = securityAccessDefinitionType;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessDefinitionType.ToLocalizedName().FormatWith (commandActivity.DisplayName, setSecurityAccessType));
      }
    }

    private static void SetSecurityAccessDefinitionOrigin (
        CommandActivity commandActivity,
        string setSecurityAccessDefinitionOrigin,
        BaseSecuredObjectWithOSS target)
    {
      try
      {
        var securityAccessDefinitionOrigin = SecurityAccessDefinitionOriginType.Values.GetValueInfoByID (setSecurityAccessDefinitionOrigin).Value;
        target.SecurityAccessDefinitionOrigin = securityAccessDefinitionOrigin;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessDefinitionOrigin.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, setSecurityAccessDefinitionOrigin));
      }
    }

    private static void AddUserOrGroup (CommandActivity commandActivity, SetOssParameters parameters, BaseSecuredObjectWithOSS target)
    {
      var access = parameters.Access;
      if (string.IsNullOrEmpty (access))
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.").WithUserMessage (
            ModuleLocalization.ParameterNullError.ToLocalizedName().FormatWith (commandActivity.DisplayName, nameof(access)));
      }

      SecurityAccessModeType accessModeType;
      try
      {
        accessModeType = SecurityAccessModeType.Values.GetValueInfoByID (access).Value;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessModeType.ToLocalizedName().FormatWith (commandActivity.DisplayName, access));
      }

      var userOrGroup = parameters.AddUserOrGroup.Invoke<TenantBoundOrganizationalStructureObject> (commandActivity, commandActivity.WorkItem);
      if (userOrGroup == null)
      {
        throw new ActivityCommandException ($"AddGroupExpression returned null in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.ExpressionEmptyResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.AddUserOrGroup)));
      }
      
      var ossEntry = ObjectSpecificSecurityEntry.NewObject (target.ObjectSpecificSecuritySetting);
      ossEntry.Access = accessModeType;
      ossEntry.Member = userOrGroup;
    }

    private static void RemoveUserOrGroup (CommandActivity commandActivity, SetOssParameters parameters, BaseSecuredObjectWithOSS target)
    {
      var access = parameters.Access;
      if (string.IsNullOrEmpty (access))
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.").WithUserMessage (
            ModuleLocalization.ParameterNullError.ToLocalizedName().FormatWith (commandActivity.DisplayName, nameof(access)));
      }

      SecurityAccessModeType accessModeType;
      try
      {
        accessModeType = SecurityAccessModeType.Values.GetValueInfoByID (access).Value;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessModeType.ToLocalizedName().FormatWith (commandActivity.DisplayName, access));
      }

      var userOrGroup = parameters.RemoveUserOrGroup.Invoke<TenantBoundOrganizationalStructureObject> (commandActivity, commandActivity.WorkItem);
      if (userOrGroup == null)
      {
        throw new ActivityCommandException ($"RemoveGroupExpression returned null in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.ExpressionEmptyResultError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.RemoveUserOrGroup)));
      }

      var entry = target.VisibleObjectSpecificSecurityEntries.FirstOrDefault (e => e.Member == userOrGroup && e.Access == accessModeType);
      if (entry == null)
      {
        throw new ActivityCommandException ($"No matching OSS entry found in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.OssEntryNotFound.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, userOrGroup.DisplayName, accessModeType));
      }

      if (!target.VisibleObjectSpecificSecurityEntries.Remove (entry))
      {
        throw new ActivityCommandException ($"Removing OSS Entry failed in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.OssEntryNotRemoved.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, userOrGroup.DisplayName, accessModeType));
      }

      entry.Delete();
    }

    private static void SetWorkflowAccessUserCurrentlyType (
        CommandActivity commandActivity,
        string setWorkflowAccessUserCurrentlyType,
        BaseSecuredObjectWithOSS target)
    {
      try
      {
        var securityAccessModeType = SecurityAccessModeType.Values.GetValueInfoByID (setWorkflowAccessUserCurrentlyType).Value;
        target.VisibleCurrentlyParticipatingInProcessAccess = securityAccessModeType;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessModeType.ToLocalizedName().FormatWith (commandActivity.DisplayName, setWorkflowAccessUserCurrentlyType));
      }
    }

    private static void SetWorkflowAccessGroupCurrentlyType (
        CommandActivity commandActivity,
        string setWorkflowAccessGroupCurrentlyType,
        BaseSecuredObjectWithOSS target)
    {
      try
      {
        var securityAccessModeType = SecurityAccessModeType.Values.GetValueInfoByID (setWorkflowAccessGroupCurrentlyType).Value;
        target.VisibleRoleCurrentlyParticipatingInProcessAccess = securityAccessModeType;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessModeType.ToLocalizedName().FormatWith (commandActivity.DisplayName, setWorkflowAccessGroupCurrentlyType));
      }
    }

    private static void SetWorkflowAccessUserPreviouslyType (
        CommandActivity commandActivity,
        string setWorkflowAccessUserPreviouslyType,
        BaseSecuredObjectWithOSS target)
    {
      try
      {
        var securityAccessModeType = SecurityAccessModeType.Values.GetValueInfoByID (setWorkflowAccessUserPreviouslyType).Value;
        target.VisiblePreviouslyParticipatingInProcessAccess = securityAccessModeType;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessModeType.ToLocalizedName().FormatWith (commandActivity.DisplayName, setWorkflowAccessUserPreviouslyType));
      }
    }

    private static void SetWorkflowAccessGroupPreviouslyType (
        CommandActivity commandActivity,
        string setWorkflowAccessGroupPreviouslyType,
        BaseSecuredObjectWithOSS target)
    {
      try
      {
        var securityAccessModeType = SecurityAccessModeType.Values.GetValueInfoByID (setWorkflowAccessGroupPreviouslyType).Value;
        target.VisibleRolePreviouslyParticipatingInProcessAccess = securityAccessModeType;
      }
      catch (KeyNotFoundException ex)
      {
        throw new ActivityCommandException ($"Execution of CommandActivity '{commandActivity.ID}' failed.", ex).WithUserMessage (
            ModuleLocalization.UnknownSecurityAccessModeType.ToLocalizedName().FormatWith (commandActivity.DisplayName, setWorkflowAccessGroupPreviouslyType));
      }
    }

    private BaseSecuredObjectWithOSS GetTargetFromWorkItemOrParam (CommandActivity commandActivity, SetOssParameters parameters)
    {
      var targetExpression = parameters.Target;
      BaseSecuredObjectWithOSS target;
      if (targetExpression != null)
      {
        target = targetExpression.Invoke<BaseSecuredObjectWithOSS> (commandActivity, commandActivity.WorkItem);

        if (target == null)
        {
          throw new ActivityCommandException ($"TargetExpression returned null in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.ExpressionEmptyResultError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, nameof(parameters.Target)));
        }

        if (target != commandActivity.WorkItem)
        {
          var isChild = false;

          if (commandActivity.WorkItem is File file && file.FileCases.Contains (target.ID))
          {
            isChild = true;
          }

          if (commandActivity.WorkItem is FileCase fileCase && fileCase.FileCaseContents.Any (c => c.ContentObject == target))
          {
            isChild = true;
          }

          if (commandActivity.WorkItem is FileHandlingObject fho && fho.FileContentHierarchyFlat.Contains (target.ID))
          {
            isChild = true;
          }

          if (!isChild)
          {
            throw new ActivityCommandException ($"Can not change OSS of objects which are not a child of the workitem in '{commandActivity.ID}'.")
                .WithUserMessage (ModuleLocalization.TargetIsNoChildOfWorkItem.ToLocalizedName().FormatWith (commandActivity.DisplayName));
          }
        }
      }
      else if (commandActivity.WorkItem is BaseSecuredObjectWithOSS workItem)
      {
        target = workItem;
      }
      else
      {
        throw new ActivityCommandException (
                $"Target must be of type '{typeof (BaseSecuredObjectWithOSS)}' but was '{commandActivity.WorkItem.GetPublicDomainObjectType().Name}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                ModuleLocalization.TargetTypeError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    typeof (BaseSecuredObjectWithOSS).ToLocalizedName()));
      }

      return target;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        if (!commandActivity.InstanceReadOnly)
        {
          throw new ActivityCommandException ($"Activity must be set to instance read only'{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.InstanceNotReadOnlyError.ToLocalizedName());
        }

        var parameters = ActivityCommandParser.Parse<SetOssParameters> (commandActivity);

        var properties = typeof (SetOssParameters).GetProperties();
        var count = properties.Where (property => property.Name != nameof(SetOssParameters.Target)
                                                  && property.Name != nameof(SetOssParameters.Access)
                                                  && property.Name != nameof(SetOssParameters.CompleteOnError)
                                                  && property.Name != nameof(SetOssParameters.ErrorSuffix)
                                                  && property.Name != nameof(SetOssParameters.DisableNotification)
                                                  && property.Name != nameof(SetOssParameters.NotificationRecipient))
            .Count (property => property.GetValue (parameters) != null);

        if (count == 0)
        {
          throw new ActivityCommandException ($"Error while parsing activity'{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.InvalidParameterError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }

        if (count > 1)
        {
          throw new ActivityCommandException ($"Can not use multiple functions in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.MultipleFunctionsError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }

        if ((parameters.AddUserOrGroup != null || parameters.RemoveUserOrGroup != null) && parameters.Access == null)
        {
          throw new ActivityCommandException ($"Error while parsing activity'{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.InvalidParameterError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }

        if (parameters.DisableNotification == false && parameters.NotificationRecipient == null)
        {
          throw new ActivityCommandException (
                  $"The notification recipient must not be null if the notifications are enabled in activity'{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.NotificationRecipientNullError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, nameof(parameters.NotificationRecipient)));
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