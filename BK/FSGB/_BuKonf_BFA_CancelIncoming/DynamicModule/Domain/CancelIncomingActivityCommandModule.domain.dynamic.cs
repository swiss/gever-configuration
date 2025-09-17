// Name: Eingang stornieren (CancelIncoming)
// Version: 1.0.1
// Datum: 16.11.2023
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.Linq;
using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Utilities;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Security;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  public class CancelIncomingActivityCommandModule : GeverActivityCommandModule
  {
    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Benutzer ist für diese Aktion nicht berechtigt.")]
      PermissionDeniedException,

      [De ("Befehlsaktivität \"{0}\": Das Objekt kann nicht storniert werden.")]
      CannotCancelObjectException,

      [De ("Befehlsaktivität \"{0}\": Die Aktivität darf kein aktives Aktivitätsobjekt benötigen.")]
      ActiveWorkitemException,

      [De ("Befehlsaktivität \"{0}\": Die Befehlsaktivität muss die letzte Aktivität im Prozess sein.")]
      LastActivityException,
    }

    public CancelIncomingActivityCommandModule ()
        : base ("CancelIncoming:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var toCancel = (BaseIncoming)commandActivity.WorkItem;

      if (toCancel.FileHandlingState == FileHandlingStateType.Canceled)
        return true;

      if (!AccessControlUtility.HasAccess (toCancel, FileHandlingObject.AccessTypes.Cancel))
      {
        throw new ActivityCommandException (
                $"The user does not have permissions to cancel in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.PermissionDeniedException.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName));
      }

      if (!toCancel.CanCancelObject())
      {
        throw new ActivityCommandException (
                $"Object cannot be canceled in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.CannotCancelObjectException.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName));
      }

      toCancel.CancelObject();

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var typeValidator = new WorkItemTypeValidator (typeof (Incoming));
        typeValidator.ValidateWorkItemType (commandActivity);

        if (commandActivity.RequiresActiveWorkItem)
        {
          throw new ActivityCommandException (
                  $"The Activity must not require an active workitem in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.ActiveWorkitemException.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName));
        }

        if (commandActivity.GetNextActivitiesAfter (new ActivityBase.GetNextActivitiesSpecification<ActivityBase>()).Any())
        {
          throw new ActivityCommandException (
                  $"The command activity must be the last activity in the process in activity '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.LastActivityException.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName));
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