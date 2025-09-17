// Name: Kommentar Hinzufügen (AddComment))
// Version: 1.0.0
// Datum: 07.08.2023
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Remotion.Security;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using Annotation = ActaNova.Domain.Annotation;

namespace Bukonf.Gever.Workflow
{
  public class AddCommentActivityCommandModule : GeverActivityCommandModule, ITemplateOnlyActivityCommandModule
  {
    public class AddCommentParameters : GeverActivityCommandParameters
    {
      [Required]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Comment { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nicht leer sein.")]
      ParameterNullError,

      [De ("Befehlsaktivität \"{0}\": Das Objekt darf nicht im Status \"{1}\" sein.")]
      FileHandlingStateInvalid,

      [De ("Die Aktivität muss schreibgeschützt in der Prozessinstanz sein.")]
      InstanceNotReadOnlyError,
    }

    public AddCommentActivityCommandModule ()
        : base ("AddComment:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<AddCommentParameters> (commandActivity);
      var commentExpression = parameters.Comment;

      var comment = commentExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);

      if (string.IsNullOrWhiteSpace (comment))
      {
        throw new ActivityCommandException (
                $"Parameter '{nameof(AddCommentParameters.Comment)}' must not be empty in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.ParameterNullError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(AddCommentParameters.Comment)));
      }

      using (SecurityFreeSection.Activate())
      {
        var target = (FileHandlingObject)commandActivity.WorkItem;

        if (target.FileHandlingState == FileHandlingStateType.Archived || target.FileHandlingState == FileHandlingStateType.Destroyed
                                                                       || target.FileHandlingState == FileHandlingStateType.CanceledDestroyed)
        {
          throw new ActivityCommandException (
                  $"FileHandlingState must not be '{target.FileHandlingState}' in activity '{commandActivity.ID}'")
              .WithUserMessage (ModuleLocalization.FileHandlingStateInvalid.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, target.FileHandlingState));
        }

        var annotation = Annotation.NewObject();
        annotation.Text = comment;
        annotation.AnnotatedObject = target;
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        var typeValidator = new WorkItemTypeValidator (typeof (FileHandlingObject));
        typeValidator.ValidateWorkItemType (commandActivity);

        if (!commandActivity.InstanceReadOnly)
        {
          throw new ActivityCommandException ($"Activity must be set to instance read only '{commandActivity.ID}'.")
              .WithUserMessage (ModuleLocalization.InstanceNotReadOnlyError.ToLocalizedName());
        }

        var parameters = ActivityCommandParser.Parse<AddCommentParameters> (commandActivity);
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