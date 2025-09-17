// Name: Insert Process Template
// Version: 2.0.0
// Datum: 09.10.2023
// Autor: RUBICON IT - Patrick Nanzer - patrick.nanzer@rubicon.eu; Ursprünglich LHIND - philipp.roessler@gs-ejpd.admin.ch

using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Remotion.Logging;
using Rubicon.Domain;
using Rubicon.Utilities;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using System;
using Rubicon.Utilities.Expr;

namespace PoC.TestingOnly.InsertProcessDefinitionByTemplate
{
  [LocalizationEnum]
  public enum LocalizedUserMessages
  {
    [De ("100: Das Geschäftsobjekt der Aktivität \"{0}\" hat einen unzulässigen Typ."),
     MultiLingualName ("L’objet métier de l’activité \"{0}\" doit être un dossier.", "Fr"),
     MultiLingualName ("L'oggetto business dell'attività \"{0}\" deve essere un dossier.", "It")]
    InvalidBusinessObject,

    [De ("200: Die Befehlsaktivität kann nur innerhalb eines Platzhalters für sequenzielle Aktivitäten verwendet werden."),
     MultiLingualName ("L’objet métier de l’activité \"{0}\" doit être un dossier.", "Fr"),
     MultiLingualName ("L'oggetto business dell'attività \"{0}\" deve essere un dossier.", "It")]
    InvalidParentActivity
  }

  public class GeverBundInsertProcessDefinitionByTemplate : ActivityCommandModule
  {
    private static readonly ILog s_logger = LogManager.GetLogger (typeof (GeverBundInsertProcessDefinitionByTemplate));

    public GeverBundInsertProcessDefinitionByTemplate () : base ("GEVERBund_InsertProcessDefinitionByTemplate:ActivityCommandClassificationType")
    {
    }

    public override bool Execute (CommandActivity commandActivity)
    {
      var onError = commandActivity.GetOrCreateFlowVariable ("onError_InsertProcessDefinitionByTemplate");

      try
      {
        if (commandActivity.WorkItem.ID.ClassID != "FileCase" && commandActivity.WorkItem.ID.ClassID != "File"
                                                              && commandActivity.WorkItem.ID.ClassID == "Incoming")
        {
          throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.InvalidBusinessObject
              .ToLocalizedName()
              .FormatWith (commandActivity));
        }

        SequentialActivity activityContainer;
        var processIdParam = string.Empty;

        s_logger.Error ("Full Parameter: " + commandActivity.Parameter);

        if (commandActivity.Parameter != null)
        {
          //Dynamic parameter calculation
          var workItemType = commandActivity.WorkItem != null ? commandActivity.WorkItem.GetPublicDomainObjectType() : commandActivity.EffectiveWorkItemType;

          s_logger.Error ("Before TypedParsedExpression is created.");
          var parsedExpression = ExpressionConfiguration.Default
              .Parse ("(activity, workItem) => " + commandActivity.Parameter,
                  typeof (CommandActivity),
                  workItemType).WithExpectedReturnType (typeof (string), null);

          //Log information
          s_logger.Error ("Expression IsValid: " + parsedExpression.IsValid);
          s_logger.Error ("Expression ReturnTypeIsValid: " + parsedExpression.ReturnTypeIsValid);
          s_logger.Error ("Expression String: " + parsedExpression.ExpressionString);
          if (parsedExpression.Exception != null)
            s_logger.Error ("Expression Exeception: " + parsedExpression.Exception);
          s_logger.Error ("Before invoke.");

          processIdParam = parsedExpression.Invoke<string> (commandActivity,
              commandActivity.WorkItem);

          if (processIdParam == null)
            throw new ActivityCommandException (
                    $"TargetExpression returned null in activity '{commandActivity.ID}'.",
                    null)
                .WithUserMessage ("Parameter konnte nicht aufgelöst werden.");

          s_logger.Error ("Parameter: " + processIdParam);
        }
        else
        {
          onError.SetValue ("100: Invalid Argument - Ungültiger Parameter");
        }

        var processTemplate = new ReferenceHandle<Template> (processIdParam).GetObject();

        if (commandActivity.GetParentActivity().ID.ClassID == "SequentialActivity")
        {
          s_logger.Error ("Sequential Activity Check - OK");

          activityContainer = (SequentialActivity)commandActivity.GetParentActivity();
        }
        else
          throw new ActivityCommandException ("").WithUserMessage (LocalizedUserMessages.InvalidParentActivity
              .ToLocalizedName()
              .FormatWith (commandActivity));

        var newActivity = processTemplate.GetInstance (true);

        var compositeActivity = activityContainer.GetParentActivity();
        compositeActivity.AppendChildActivity (activityContainer, newActivity);

        onError.SetValue ("0");
      }
      catch (Exception ex)
      {
        s_logger.Error ("Error Handling - Catch Part");
        onError.SetValue ("-1: Someting went wrong");

        s_logger.Error (ex.Message);
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      return commandActivity.Parameter != null ? null : "Bei der Aktivitaet muss ein Parameter als LINQ-Expression oder mit der Referenz-ID der einzufügenden Prozessvorlage übergeben werden.";
    }
  }
}