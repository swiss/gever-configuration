// Name: Fachdaten setzen mit Target (SetSpecialdataProperty)
// Version: 1.0.1
// Datum: 10.04.2025
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ActaNova.Domain;
using ActaNova.Domain.Specialdata;
using ActaNova.Domain.Specialdata.Values;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Remotion.ObjectBinding;
using Rubicon.Data.DomainObjects.Transport;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.ObjectBinding.Utilities.ObjectProxy;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  [VersionInfo("1.0.1", "10.04.2025", "RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu")]
  public class SetSpecialdataPropertyActivityCommandModule : GeverActivityCommandModule
  {
    [AttributeUsage(AttributeTargets.Class, Inherited = false)]
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
    public enum Localization
    {
      [De ("Aktivität \"{0}\": Der Ausdruck \"{1}\" kann nicht verarbeitet werden.")]
      InvalidExpression,

      [De ("Aktivität \"{0}\": Der Wert kann der Fachdaten-Eigenschaft \"{1}\" nicht zugewiesen werden, weil die Fachdaten-Eigenschaft nicht existiert.")]
      InvalidSpecialdataPropertyIdentifier,

      [De (
          "Aktivität \"{0}\": Der Wert kann der Fachdaten-Eigenschaft \"{1}\" nicht zugewiesen werden, weil beim Parsen des Ausdrucks ein Fehler aufgetreten ist.")]
      ParseError,

      [De ("Aktivität \"{0}\": Beim Parsen des Ausdrucks \"{1}\" ist ein Fehler aufgetreten.")]
      GenericParseError,

      [De (
          "Aktivität \"{0}\": Der Wert kann der Fachdaten-Eigenschaft \"{1}\" nicht zugewiesen werden, weil bei der Berechnung des Ausdrucks ein Fehler aufgetreten ist.")]
      EvaluationError,

      [De (
          "Aktivität \"{0}\": Der Wert kann der Fachdaten-Eigenschaft \"{1}\" nicht zugewiesen werden, weil die Fachdaten-Eigenschaft für das Aktivitätsobjekt nicht verfügbar ist.")]
      SpecialdataPropertyNotAvailable,

      [De (
          "Aktivität \"{0}\": Der Wert kann der Fachdaten-Eigenschaft \"{1}\" nicht zugewiesen werden, weil die Fachdaten-Eigenschaft in einem Fachdaten-Aggregat ist. Sie müssen die Fachdaten-Eigenschaft indirekt setzen.")]
      InvalidDirectAggregatePropertyAccess,

      [De (
          "Aktivität \"{0}\": Der Wert kann der Fachdaten-Eigenschaft \"{1}\" nicht zugewiesen werden, weil die Fachdaten-Eigenschaft nicht in einem Fachdaten-Aggregat ist. Sie müssen die Fachdaten-Eigenschaft direkt setzen.")]
      InvalidIndirectNonAggregatePropertyAccess,

      [De ("Befehlsaktivität \"{0}\": Das Zieldossier ist ungültig.")]
      TargetInvalidError,

      [De("Befehlsaktivität \"{0}\": Version: \"{1}\" - Datum: \"{2}\" - Autor \"{3}\"")]
      ShowVersion
    }

    public SetSpecialdataPropertyActivityCommandModule ()
        : base ("SetSpecialdataProperty:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var workItem = commandActivity.WorkItem;
      if (!workItem.CanEdit (true))
        throw new ActivityCommandException ($"Workitem '{workItem.DisplayName}' is not editable.");

      var parameters = commandActivity.Parameter;
      var specialDataValueTarget = ExtractSpecialDataValueTarget (ref parameters, commandActivity) ?? workItem;
      if (!ValidateTarget (commandActivity, specialDataValueTarget))
      {
        throw new ActivityCommandException (
                $"The Target '{specialDataValueTarget.GetDisplayName()}' is invalid in activity '{commandActivity.ID}'")
            .WithUserMessage (Localization.TargetInvalidError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName));
      }

      var executionContexts = GetExecutionContexts (parameters, specialDataValueTarget, commandActivity);

      foreach (var executionContext in executionContexts)
      {
        var invocationResult = GetInvocationResult (commandActivity, executionContext);

        try
        {
          IBusinessObject specialdataValueTarget = executionContext.SpecialDataValueTarget;
          var specialdataValueTargetPropertyIdentifier = executionContext.OuterSpecialdataPropertyIdentifier;

          if (!string.IsNullOrEmpty (executionContext.InnerSpecialdataPropertyIdentifier))
          {
            if (!executionContext.OuterSpecialdataPropertyIdentifier.EndsWith (VirtualDomainObjectProxyPropertyInformation.ResolvedPropertyNameSuffix))
              specialdataValueTargetPropertyIdentifier =
                  $"{specialdataValueTargetPropertyIdentifier}{VirtualDomainObjectProxyPropertyInformation.ResolvedPropertyNameSuffix}";

            specialdataValueTarget = specialdataValueTarget.GetProperty (specialdataValueTargetPropertyIdentifier) as IBusinessObject;
            specialdataValueTargetPropertyIdentifier = executionContext.InnerSpecialdataPropertyIdentifier;
          }

          using (new SpecialdataIgnoreReadOnlySection())
          {
            specialdataValueTarget.SetProperty (specialdataValueTargetPropertyIdentifier, invocationResult);
          }
        }
        catch (Exception ex)
        {
          if (ex is ActivityCommandException)
            throw;

          throw new ActivityCommandException (
                  $"Specialdata property '{executionContext.CompletePropertyIdentifier}' is not availabe on workitem '{commandActivity.WorkItem}'.",
                  ex)
              .WithUserMessage (Localization.SpecialdataPropertyNotAvailable.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, executionContext.CompletePropertyIdentifier));
        }
      }

      return true;
    }

    private object GetInvocationResult (CommandActivity commandActivity, ExecutionContext executionContext)
    {
      try
      {
        var invocationResult = executionContext.Expression.Invoke (commandActivity, commandActivity.WorkItem);
        return invocationResult;
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException ($"'{ex}'.", ex).WithUserMessage (Localization.EvaluationError.ToLocalizedName()
            .FormatWith (executionContext.ActivityDisplayName, executionContext.CompletePropertyIdentifier));
      }
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        // if no exception is command activity is valid up so far
        // ReSharper disable once ReturnValueOfPureMethodIsNotUsed
        var parameters = commandActivity.Parameter;
        ShowVersion(commandActivity);
        _ = ExtractSpecialDataValueTarget (ref parameters, commandActivity);
        _ = GetExecutionContexts (parameters, null, commandActivity);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    private void ShowVersion (CommandActivity commandActivity)
    {
      foreach (var line in Regex.Split (commandActivity.Parameter, "\r\n|\r|\n"))
      {
        if (line.Trim().ToLower().Replace (" ", "").StartsWith ("showversion"))
        {
          var attr = GetType().GetCustomAttribute<VersionInfoAttribute>();
          if (attr != null)
          {
            throw new ActivityCommandException(
                    $"CommandActivity \"{commandActivity.DisplayName}\": Version: \"{attr.Version}\" - Date: \"{attr.Date}\" - Author \"{attr.Author}\"")
                .WithUserMessage(Localization
                    .ShowVersion.ToLocalizedName().FormatWith(
                        commandActivity.DisplayName,
                        attr.Version,
                        attr.Date,
                        attr.Author));
          }
        }
      }
    }

    private IEnumerable<ExecutionContext> GetExecutionContexts (string parameter, IBusinessObject specialDataValueTarget, CommandActivity commandActivity)
    {
      if (string.IsNullOrEmpty (parameter))
        yield break;

      foreach (var commandLine in Regex.Split (parameter, "\r\n|\r|\n"))
      {
        var trimmedLine = commandLine.Trim();
        if (trimmedLine.IsNullOrEmpty() || trimmedLine.StartsWith ("//"))
          continue;

        var executionInfo = PrepareForExcecution (trimmedLine, specialDataValueTarget, commandActivity);
        yield return executionInfo;
      }
    }

    private IBusinessObject ExtractSpecialDataValueTarget (ref string parameter, CommandActivity commandActivity)
    {
      var firstLine = Regex.Split (parameter, "\r\n|\r|\n").First().Trim();

      var commandParts = firstLine.Split (new[] { '=' }, 2);
      if (commandParts.Length != 2)
      {
        throw new ActivityCommandException ($"Parameter line '{firstLine}' can not be processed.")
            .WithUserMessage (Localization.InvalidExpression.ToLocalizedName().FormatWith (commandActivity.DisplayName, firstLine));
      }

      if (commandParts[0].ToLower().Trim() != "target")
        return null;

      parameter = parameter.Replace (firstLine, string.Empty);

      if (commandActivity.ActivityState == ActivityStateType.Template || commandActivity.ActivityState == ActivityStateType.Initialized)
        return null;

      try
      {
        var expression = ParseExpression (commandParts[1], typeof (IBusinessObject), commandActivity);
        var specialDataValueTarget = expression.Invoke<IBusinessObject> (commandActivity, commandActivity.WorkItem);
        return specialDataValueTarget;
      }
      catch (Exception ex)
      {
        var userMessage = ex.GetUserMessage();
        if (string.IsNullOrEmpty (userMessage))
          userMessage = Localization.GenericParseError.ToLocalizedName().FormatWith (commandActivity.DisplayName, firstLine);

        if (ex is TargetInvocationException && ex.InnerException != null)
          ex = ex.InnerException;

        if (ex is ActivityCommandException)
          throw new ActivityCommandException ($"'{ex}'.", ex).WithUserMessage (Localization.GenericParseError.ToLocalizedName()
              .FormatWith (commandActivity.DisplayName, firstLine));

        throw new ActivityCommandException (ex.Message, ex).WithUserMessage (userMessage);
      }
    }

    private bool ValidateTarget (CommandActivity commandActivity, IBusinessObject specialDataValueTarget)
    {
      if (specialDataValueTarget.Equals (commandActivity.WorkItem))
        return true;

      var workItemType = commandActivity.WorkItem?.GetPublicDomainObjectType().Name ?? commandActivity.EffectiveWorkItemType.Name;

      if (workItemType == "FileCase")
      {
        var workItem = (FileCase)commandActivity.WorkItem;
        var workItemParent = !workItem.HasParent ? workItem : workItem.ParentObject;

        if (workItemParent.Equals (specialDataValueTarget))
          return true;

        return workItem.BaseFile.Documents.Contains (specialDataValueTarget);
      }

      var typedWorkItem = (FileHandlingObject)commandActivity.WorkItem;
      var typedSpecialDataValueTarget = (FileHandlingObject)specialDataValueTarget;
      var workItemParentObject = !typedWorkItem.HasParent ? typedWorkItem : typedWorkItem.ParentObject;

      if (workItemParentObject.Equals (specialDataValueTarget))
        return true;

      while (typedSpecialDataValueTarget.HasParent)
      {
        if (typedSpecialDataValueTarget.ParentObject.Equals (typedWorkItem))
          return true;

        typedSpecialDataValueTarget = (FileHandlingObject)typedSpecialDataValueTarget.ParentObject;
      }

      return false;
    }

    private ExecutionContext PrepareForExcecution (string trimmedCommandLine, IBusinessObject specialDataValueTarget, CommandActivity commandActivity)
    {
      var executionContext = new ExecutionContext
      {
          SpecialDataValueTarget = specialDataValueTarget,
          OriginalTrimmedCommandLine = trimmedCommandLine,
          ActivityDisplayName = commandActivity.DisplayName
      };

      var commandParts = GetCommandParts (executionContext);

      SetSpecialdataPropertyIdentifiers (executionContext, commandParts[0].Trim());
      SetExpressionString (executionContext, commandParts[1].Trim());

      ParseExpressionString (executionContext, commandActivity);

      return executionContext;
    }

    private void SetExpressionString (ExecutionContext executionContext, string expressionString)
    {
      if (expressionString.IsNullOrEmpty())
      {
        throw new ActivityCommandException ($"Parameter line '{executionContext.OriginalTrimmedCommandLine}' can not be processed.")
            .WithUserMessage (Localization.InvalidExpression.ToLocalizedName()
                .FormatWith (executionContext.ActivityDisplayName, executionContext.OriginalTrimmedCommandLine));
      }

      executionContext.ExpressionString = expressionString;
    }

    private void SetSpecialdataPropertyIdentifiers (ExecutionContext executionContext, string specialdataPropertyIdentifier)
    {
      var match = Regex.Matches (specialdataPropertyIdentifier, @"(#\w+)(\.(#w+))?");
      if (match.Count < 1 || match.Count > 2)
      {
        throw new ActivityCommandException ($"Parameter line '{executionContext.OriginalTrimmedCommandLine}' can not be processed.")
            .WithUserMessage (Localization.InvalidExpression.ToLocalizedName()
                .FormatWith (executionContext.ActivityDisplayName, executionContext.OriginalTrimmedCommandLine));
      }

      executionContext.OuterSpecialdataPropertyIdentifier = match[0].Value;
      executionContext.OuterSpecialdataPropertyType = executionContext.SpecialDataValueTargetType.GetBusinessObjectClass()
          .GetPropertyDefinition (executionContext.OuterSpecialdataPropertyIdentifier)?.PropertyType;

      if (executionContext.OuterSpecialdataPropertyType == null)
      {
        throw new ActivityCommandException ($"Identifier '{executionContext.OuterSpecialdataPropertyIdentifier}' is not a valid special data property.")
            .WithUserMessage (Localization.InvalidSpecialdataPropertyIdentifier.ToLocalizedName()
                .FormatWith (executionContext.ActivityDisplayName, executionContext.OuterSpecialdataPropertyIdentifier));
      }

      if (match.Count == 2)
      {
        if (executionContext.OuterSpecialdataPropertyType != typeof (SpecialdataAggregatePropertyValue))
        {
          throw new ActivityCommandException (
                  $"Identifier '{executionContext.OuterSpecialdataPropertyIdentifier}' referres to a Non-SpecialdataAggregateProperty which cannnot be set indirectly.")
              .WithUserMessage (Localization.InvalidIndirectNonAggregatePropertyAccess.ToLocalizedName()
                  .FormatWith (executionContext.ActivityDisplayName, executionContext.CompletePropertyIdentifier));
        }

        executionContext.InnerSpecialdataPropertyIdentifier = match[1].Value;
      }

      if (executionContext.OuterSpecialdataPropertyType == typeof (SpecialdataAggregatePropertyValue))
      {
        if (executionContext.InnerSpecialdataPropertyIdentifier.IsNullOrEmpty())
        {
          throw new ActivityCommandException (
                  $"Identifier '{executionContext.OuterSpecialdataPropertyIdentifier}' referres to a SpecialdataAggregateProperty which cannot be set directly.")
              .WithUserMessage (Localization.InvalidDirectAggregatePropertyAccess.ToLocalizedName()
                  .FormatWith (executionContext.ActivityDisplayName, executionContext.OuterSpecialdataPropertyIdentifier));
        }

        executionContext.InnerSpecialdataPropertyType = executionContext.OuterSpecialdataPropertyType.GetBusinessObjectClass()
            .GetPropertyDefinition (executionContext.InnerSpecialdataPropertyIdentifier)?.PropertyType;
      }

      if (executionContext.InnerSpecialdataPropertyIdentifier != null && executionContext.InnerSpecialdataPropertyType == null)
      {
        throw new ActivityCommandException ($"Identifier '{executionContext.InnerSpecialdataPropertyIdentifier}' is not a valid special data property.")
            .WithUserMessage (Localization.InvalidSpecialdataPropertyIdentifier.ToLocalizedName()
                .FormatWith (executionContext.ActivityDisplayName, executionContext.CompletePropertyIdentifier));
      }
    }

    private string[] GetCommandParts (ExecutionContext executionContext)
    {
      var commandParts = executionContext.OriginalTrimmedCommandLine.Split (new[] { '=' }, 2);
      if (commandParts.Length != 2)
      {
        throw new ActivityCommandException ($"Parameter line '{executionContext.OriginalTrimmedCommandLine}' can not be processed.")
            .WithUserMessage (Localization.InvalidExpression.ToLocalizedName()
                .FormatWith (executionContext.ActivityDisplayName, executionContext.OriginalTrimmedCommandLine));
      }

      return commandParts;
    }

    private void ParseExpressionString (ExecutionContext executionContext, CommandActivity commandActivity)
    {
      var expectedReturnType = executionContext.InnerSpecialdataPropertyType ?? executionContext.OuterSpecialdataPropertyType;

      try
      {
        var parsedExpression = ParseExpression (executionContext.ExpressionString, expectedReturnType, commandActivity);

        executionContext.Expression = parsedExpression;
      }
      catch (Exception ex)
      {
        var userMessage = ex.GetUserMessage();
        if (string.IsNullOrEmpty (userMessage))
          userMessage = Localization.ParseError.ToLocalizedName()
              .FormatWith (executionContext.ActivityDisplayName, executionContext.CompletePropertyIdentifier);

        if (ex is TargetInvocationException && ex.InnerException != null)
          ex = ex.InnerException;

        if (ex is ActivityCommandException)
          throw new ActivityCommandException ($"'{ex}'.", ex).WithUserMessage (Localization.ParseError.ToLocalizedName()
              .FormatWith (executionContext.ActivityDisplayName, executionContext.CompletePropertyIdentifier));

        throw new ActivityCommandException (ex.Message, ex).WithUserMessage (userMessage);
      }
    }

    private static TypedParsedExpression ParseExpression (string expressionString, Type expectedReturnType, CommandActivity commandActivity)
    {
      var formattedExpressionString = $"(activity, workItem) => {expressionString}";
      var parsedExpression = ExpressionConfiguration.Default.Parse (formattedExpressionString,
              typeof (CommandActivity),
              commandActivity.WorkItem?.GetPublicDomainObjectType() ?? commandActivity.EffectiveWorkItemType)
          .WithExpectedReturnType (expectedReturnType);

      var parseException = parsedExpression.Exception;
      if (!parsedExpression.IsValid || parseException != null)
      {
        var typeNameForCast = expectedReturnType.Name;
        if (expectedReturnType.IsNullable())
          typeNameForCast = $"{expectedReturnType.NullableUnderlying().FullName}?";

        formattedExpressionString = $"(activity, workItem) => {typeNameForCast}({expressionString})";
        parsedExpression = ExpressionConfiguration.Default.Parse (formattedExpressionString,
                typeof (CommandActivity),
                commandActivity.WorkItem?.GetPublicDomainObjectType() ?? commandActivity.EffectiveWorkItemType)
            .WithExpectedReturnType (expectedReturnType);
      }

      parseException = parsedExpression.Exception;
      if (!parsedExpression.IsValid || parseException != null)
      {
        throw new ActivityCommandException (
                parseException?.Message ?? $"Error while parsing expression '{expressionString}'.",
                parseException)
            .WithUserMessage (Localization.GenericParseError.ToLocalizedName().FormatWith (commandActivity.DisplayName,
                expressionString));
      }

      return parsedExpression;
    }

    private class ExecutionContext
    {
      public string ActivityDisplayName { get; set; }

      public string CompletePropertyIdentifier => InnerSpecialdataPropertyIdentifier.IsNullOrEmpty()
          ? OuterSpecialdataPropertyIdentifier
          : $"{OuterSpecialdataPropertyIdentifier}.{InnerSpecialdataPropertyIdentifier}";

      public string OuterSpecialdataPropertyIdentifier { get; set; }
      public string InnerSpecialdataPropertyIdentifier { get; set; }
      public Type OuterSpecialdataPropertyType { get; set; }
      public Type InnerSpecialdataPropertyType { get; set; }
      public Type SpecialDataValueTargetType => SpecialDataValueTarget.GetType();
      public IBusinessObject SpecialDataValueTarget { get; set; }
      public TypedParsedExpression Expression { get; set; }
      public string ExpressionString { get; set; }
      public string OriginalTrimmedCommandLine { get; set; }
    }
  }
}