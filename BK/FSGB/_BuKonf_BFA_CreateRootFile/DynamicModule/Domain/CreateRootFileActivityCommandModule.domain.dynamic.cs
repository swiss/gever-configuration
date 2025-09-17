// Name: Root-Dossier erstellen (createRootFile)
// Version: 1.0.2
// Datum: 06.11.2023
// Autor: RUBICON IT - Delia Jacobs - delia.jacobs@rubicon.eu

using System;
using System.Collections.Generic;
using System.Linq;
using ActaNova.Domain;
using ActaNova.Domain.Classifications;
using ActaNova.Domain.Utilities;
using ActaNova.Domain.Workflow;
using Remotion.Data.DomainObjects.Queries;
using Remotion.Globalization;
using Remotion.SecurityManager.Domain.OrganizationalStructure;
using Rubicon.Domain;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Mixins;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.Multilingual;
using Rubicon.Multilingual.Extensions;
using Rubicon.Multilingual.Shared;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Inheritable;
using Rubicon.Workflow.Domain;
using Rubicon.Gever.Bund.AccessDefinitions.Domain;
using Rubicon.Utilities.Security;

namespace Bukonf.Gever.Workflow
{
  public class CreateRootFileActivityCommandModule : GeverActivityCommandModule, ITemplateOnlyActivityCommandModule
  {
    protected string CreatedObjectFlowVariableName => "createdFileObject";

    [LocalizationEnum]
    private new enum Localization
    {
      [De ("Befehlsaktivität \"{0}\": Der Ausdruck im Parameter \"Target\" liefert kein Ergebnis.")]
      TargetExpressionEmptyResultError,

      [De ("Befehlsaktivität \"{0}\": Die Ordnungsposition muss eine Rubrik sein.")]
      TargetLeafError,

      [De ("Befehlsaktivität \"{0}\": Es konnte keine Vorlage mit der Referenz-ID \"{1}\" gefunden werden.")]
      TemplateReferenceIDNotFoundError,

      [De ("Befehlsaktivität \"{0}\": Es konnte keine Vorlage mit dem Namen \"{1}\" gefunden werden.")]
      TemplateDisplayNameNotFoundError,

      [De ("Befehlsaktivität \"{0}\": Es wurde mehr als eine Vorlage mit dem Namen \"{1}\" gefunden.")]
      TemplateDisplayNameMultipleResultsError,

      [De ("Die Aktivität muss schreibgeschützt in der Prozessinstanz sein.")]
      InstanceNotReadOnlyError,

      [De ("Befehlsaktivität \"{0}\": Dossiertyp \"{1}\" ist in der Rubrik nicht verfügbar.")]
      FileTypeNotAvailableError,

      [De ("Der Prozess \"{0}\" hat ein neues Root-Dossier erstellt!")]
      RootFileCreatedMesssage
    }

    public class RootFileParameters : GeverActivityCommandParameters
    {
      [Required]
      [ExpectedExpressionType (typeof (SubjectArea))]
      public TypedParsedExpression Target { get; set; }

      [RequiredOneOfMany (new[] { nameof(TemplateReferenceID), nameof(TemplateDisplayName), nameof(Template) })]
      [EitherOr (new[] { nameof(TemplateReferenceID), nameof(TemplateDisplayName), nameof(Template) })]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression TemplateReferenceID { get; protected set; }

      [RequiredOneOfMany (new[] { nameof(TemplateReferenceID), nameof(TemplateDisplayName), nameof(Template) })]
      [EitherOr (new[] { nameof(TemplateReferenceID), nameof(TemplateDisplayName), nameof(Template) })]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression TemplateDisplayName { get; protected set; }

      [RequiredOneOfMany (new[] { nameof(TemplateReferenceID), nameof(TemplateDisplayName), nameof(Template) })]
      [EitherOr (new[] { nameof(TemplateReferenceID), nameof(TemplateDisplayName), nameof(Template) })]
      [ExpectedExpressionType (typeof (FileTemplate))]
      public TypedParsedExpression Template { get; set; }
    }

    public CreateRootFileActivityCommandModule ()
        : base ("CreateRootFile:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<RootFileParameters> (commandActivity);

      var templateReferenceIDExpression = parameters.TemplateReferenceID;
      var templateDisplayNameExpression = parameters.TemplateDisplayName;
      var templateExpression = parameters.Template;

      var templateReferenceID = templateReferenceIDExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var templateDisplayName = templateDisplayNameExpression?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      var templateParameter = templateExpression?.Invoke<FileTemplate> (commandActivity, commandActivity.WorkItem);

      var template = templateParameter ?? GetTemplate (commandActivity, templateReferenceID, templateDisplayName);

      var targetExpression = parameters.Target;
      var targetSubjectArea = targetExpression?.Invoke<SubjectArea> (commandActivity, commandActivity.WorkItem);

      if (targetSubjectArea == null)
      {
        throw new ActivityCommandException ($"TargetExpression returned null in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.TargetExpressionEmptyResultError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }

      if (!targetSubjectArea.IsValidAndActive() || !targetSubjectArea.CanBeSelected || targetSubjectArea.SubSubjectAreas.Count > 0)
      {
        throw new ActivityCommandException (
                $"Target Subject Area must not have sub Subject Areas in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.TargetLeafError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }

      AssertNumeratorReferenceNumberCreationMode (commandActivity, targetSubjectArea);

      if (!targetSubjectArea.CalculatedSubjectAreaFileTypeLinks.Any (t =>
              t.FileType == template.FileType && t.AvailableFor == SubjectAreaFileTypeLink.SubjectAreaFileTypeLinkType.File))
      {
        throw new ActivityCommandException ($"File type is not available for this Subject Area in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.FileTypeNotAvailableError.ToLocalizedName().FormatWith (commandActivity.DisplayName, template.FileType.DisplayName));
      }

      var file = File.NewObject();

      var templateArguments = new FileFromTemplateArguments (file)
      {
          SubjectArea = targetSubjectArea,
          FileTemplate = template,
          ApplicationContext = targetSubjectArea.TopLevelSubjectArea.ApplicationContext
      };
      file.ApplyTemplateArguments (templateArguments);

      if (file.GetMultilingualValue (o => o.Title).IsNullOrEmpty())
      {
        foreach (var culture in MultilingualConfigurationSection.Current.AllDataCultureInfos)
        {
          var value = template.GetMultilingualValue (o => o.Name, culture);
          if (!string.IsNullOrEmpty (value))
            file.SetMultilingualValue (o => ((File)o).Title, value, culture);
        }
      }

      var createdObjectFlowVariable = commandActivity.GetOrCreateFlowVariable (CreatedObjectFlowVariableName);
      createdObjectFlowVariable.SetValue (file);

      var recipients = new List<TenantGroup>();

      var registration = ((IGeverResponsibleRegistrationServiceMixin)targetSubjectArea.EffectiveLeadingGroup?.ActualGroup)?.ResponsibleRegistrationServiceGroup
          ?.AsTenantGroup();

      if (registration != null)
      {
        recipients.Add (registration);
      }
      else
      {
        var groupsWithRegistration = QueryFactory.CreateLinqQuery<Group>()
            .Where (g => g.Tenant.ID == TenantAwareSecurityManagerPrincipal.Current.Tenant.ID)
            .Where (g => g.Roles.Any (r => r.Position.ToHasReferenceIDMixin().ReferenceID == Constants.PositionRegistrationReferenceID));

        groupsWithRegistration.AsEnumerable().ForEach (g => recipients.Add (g.AsTenantGroup()));
      }

      foreach (var recipient in recipients)
      {
        SendThroughWorkflowModule.Value (
            file,
            recipient,
            Localization.RootFileCreatedMesssage.ToLocalizedName().FormatWith (((ProcessActivity)commandActivity.RootActivity).DisplayName),
            true,
            null);
      }

      return true;
    }

    protected FileTemplate GetTemplate (
        CommandActivity commandActivity,
        string templateReferenceID,
        string templateDisplayName)
    {
      if (templateReferenceID != null)
      {
        var template = new ReferenceHandle<FileTemplate> (templateReferenceID).TryGetObject();

        if (template == null)
        {
          throw new ActivityCommandException (
                  $"Template with ReferenceID '{templateReferenceID}' not found in activity '{commandActivity.ID}'.")
              .WithUserMessage (
                  Localization.TemplateReferenceIDNotFoundError.ToLocalizedName().FormatWith (
                      commandActivity.DisplayName,
                      templateReferenceID));
        }

        return template;
      }

      var templates = QueryFactory.CreateLinqQuery<FileTemplate>()
          // ReSharper disable once SuspiciousTypeConversion.Global
          .Where (o => ((IPersistedDisplayNameDomainObjectMixin)o).PersistedDisplayName == templateDisplayName)
          .ToList();

      if (templates.Count == 0)
      {
        throw new ActivityCommandException (
                $"Template with DisplayName '{templateDisplayName}' not found in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                Localization.TemplateDisplayNameNotFoundError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    templateDisplayName));
      }

      if (templates.Count > 1)
      {
        throw new ActivityCommandException (
                $"More than one Template with DisplayName '{templateDisplayName}' found in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                Localization.TemplateDisplayNameMultipleResultsError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    templateDisplayName));
      }

      return templates.First();
    }

    private void AssertNumeratorReferenceNumberCreationMode (CommandActivity commandActivity, SubjectArea subjectArea)
    {
      if (subjectArea.AsInheritable().EffectiveValue (s => s.ReferenceNumberCreationModeFile) != ReferenceNumberCreationModeType.Values.Numerator())
      {
        throw new ActivityCommandException (
                $"'{subjectArea.ID}' must have ReferenceNumberCreationMode 'Numerator' in activity '{commandActivity.ID}'.")
            .WithUserMessage (
                CreateFileHandlingObjectTemplateActivityCommandModule.Localization.NumeratorReferenceNumberCreationModeError.ToLocalizedName().FormatWith (
                    commandActivity.DisplayName,
                    ReferenceNumberCreationModeType.Values.Numerator().ToLocalizedName()));
      }
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        if (!commandActivity.InstanceReadOnly)
        {
          throw new ActivityCommandException ($"Activity must be set to instance read only'{commandActivity.ID}'.")
              .WithUserMessage (Localization.InstanceNotReadOnlyError.ToLocalizedName());
        }

        var parameters = ActivityCommandParser.Parse<RootFileParameters> (commandActivity);
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