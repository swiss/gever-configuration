// Name: Dossiereigenschaften setzen (SetFile)
// Version: 1.0.1
// Datum: 11.09.2025
// Autor: RUBICON IT - Claudia Fleck (vorher: Delia Jacobs) - claudia.fleck@rubicon.eu

using System;
using System.Linq;
using System.Reflection;
using ActaNova.Domain;
using ActaNova.Domain.Classifications;
using ActaNova.Domain.Workflow;
using JetBrains.Annotations;
using Remotion.Globalization;
using Rubicon.Domain.Interfaces;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Classifications;
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.Mixins;
using Rubicon.Gever.Bund.DocumentEncryption.Core;
using Rubicon.Gever.Bund.DocumentEncryption.Core.Classifications;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;

namespace Bukonf.Gever.Workflow
{
  [UsedImplicitly]
  [VersionInfo ("1.0.1", "11.09.2025", "RUBICON IT - Claudia Fleck (vorher: Delia Jacobs) - claudia.fleck@rubicon.eu")]
  public class SetFileActivityCommandModule : GeverActivityCommandModule
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
      [De ("Befehlsaktivität \"{0}\": Das Aktivitätsobjekt muss ein Dossier sein.")]
      InvalidWorkItemType,

      [De ("Befehlsaktivität \"{0}\": Der File Type \"{1}\" ist für diese Ordnungsposition nicht verfügbar.")]
      FileTypeNotAvailableForSubjectArea,

      [De ("Befehlsaktivität \"{0}\": Das dossier \"{1}\" muss im Status \"In Bearbeitung\" sein.")]
      FileNotInStateWorkError,

      [De ("Befehlsaktivität \"{0}\": Die maximale Verschachtelungstiefe muss grösser als 1 sein.")]
      MaxNestingLevelLessThanOneError,

      [De ("Befehlsaktivität \"{0}\": Der Wert \"{1}\" für den Parameter \"{2}\" ist inaktiv oder nicht mehr gültig.")]
      ValueInactiveOrInvalidError,

      [De (
          "Befehlsaktivität \"{0}\": Der Parameter \"{1}\" ist nicht im gültigen Wertebereich. Gültiger Wertebereich: {2} - {3}")]
      ValueNotInValidRangeError,

      [De ("Befehlsaktivität \"{0}\": Das Dossier \"{1}\" kann vom aktuellen Benutzer nicht bearbeitet werden.")]
      CannotEditFileError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" darf nur auf einem Root Dossier gesetzt werden.")]
      FileIsNotARootFileError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" ist schreibgeschützt und darf nicht editiert werden.")]
      PropertyIsReadOnlyError,

      [De ("Befehlsaktivität \"{0}\": Der Parameter \"{1}\" konnte nicht geparsed werden.")]
      DateParsingError,

      [De ("Befehlsaktivität \"{0}\": Version: \"{1}\" - Datum: \"{2}\" - Autor \"{3}\"")]
      ShowVersion,
    }

    public class SetFileParameters : GeverActivityCommandParameters
    {
      [ExpectedExpressionType (typeof (FileTypeClassificationType))]
      public TypedParsedExpression FileType { get; set; }

      [ExpectedExpressionType (typeof (DateTime))]
      public TypedParsedExpression OpeningDate { get; set; }

      [ExpectedExpressionType (typeof (DateTime))]
      public TypedParsedExpression ExternalDate { get; set; }

      [ExpectedExpressionType (typeof (ReferenceNumberCreationModeType))]
      public TypedParsedExpression ReferenceNumberCreationModeSubFile { get; set; }

      [ExpectedExpressionType (typeof (TenantGroup))]
      public TypedParsedExpression LeadingGroup { get; set; }

      [ExpectedExpressionType (typeof (TenantUser))]
      public TypedParsedExpression LeadingUser { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression Remark { get; set; }

      [ExpectedExpressionType (typeof (int))]
      public TypedParsedExpression MaxNestingLevel { get; set; }

      [ExpectedExpressionType (typeof (DataProtectionClassificationType))]
      public TypedParsedExpression DataProtection { get; set; }

      [ExpectedExpressionType (typeof (ClassificationCategoryType))]
      public TypedParsedExpression ClassificationCategory { get; set; }

      [ExpectedExpressionType (typeof (PublicDisclosureClassificationType))]
      public TypedParsedExpression PublicDisclosureIntention { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression PublicDisclosureReason { get; set; }

      [ExpectedExpressionType (typeof (RetentionPeriodCategoryClassificationType))]
      public TypedParsedExpression RetentionPeriodCategory { get; set; }

      [ExpectedExpressionType (typeof (int))]
      public TypedParsedExpression RetentionPeriodYears { get; set; }

      [ExpectedExpressionType (typeof (DeletionModeClassificationType))]
      public TypedParsedExpression DeletionMode { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression ArchivalValueReason { get; set; }

      [ExpectedExpressionType (typeof (int))]
      public TypedParsedExpression DeletionDelayYears { get; set; }

      [ExpectedExpressionType (typeof (int))]
      public TypedParsedExpression DeletionBackupYears { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression IsPaperFile { get; set; }

      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression PaperFileFilingLocation { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression ShowVersion { get; set; }
    }

    public SetFileActivityCommandModule ()
        : base ("SetFile:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      ValidateBasedataCatalogValues (commandActivity);

      var parameters = ActivityCommandParser.Parse<SetFileParameters> (commandActivity);
      var file = (File)commandActivity.WorkItem;

      if (!file.CanEdit (ignoreSuspendable: true))
      {
        throw new ActivityCommandException (
                $"User cannot Edit File '{file.DisplayName}' in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.CannotEditFileError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, file.DisplayName));
      }

      if (file.FileHandlingState != FileHandlingStateType.Work)
      {
        throw new ActivityCommandException (
                $"File'{file.DisplayName}' must have state 'Work' in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.FileNotInStateWorkError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, file.DisplayName));
      }

      SetProperty<DateTime> (commandActivity, nameof(File.OpeningDate), parameters.OpeningDate, (f, v) => f.OpeningDate = v);
      SetProperty<DateTime> (commandActivity, nameof(File.ExternalDate), parameters.ExternalDate, (f, v) => f.ExternalDate = v);
      SetProperty<ReferenceNumberCreationModeType> (commandActivity,
          nameof(File.ReferenceNumberCreationModeSubFile),
          parameters.ReferenceNumberCreationModeSubFile,
          (f, v) => f.ReferenceNumberCreationModeSubFile = v);
      SetProperty<TenantGroup> (commandActivity, nameof(File.LeadingGroup), parameters.LeadingGroup, (f, v) => f.LeadingGroup = v);
      SetProperty<TenantUser> (commandActivity, nameof(File.LeadingUser), parameters.LeadingUser, (f, v) => f.LeadingUser = v);
      SetProperty<string> (commandActivity, nameof(File.Remark), parameters.Remark, (f, v) => f.Remark = v);
      SetProperty<DataProtectionClassificationType> (commandActivity,
          nameof(IDataProtectionAware.DataProtection),
          parameters.DataProtection,
          (f, v) => f.ToDataProtectionAware().DataProtection = v);
      SetProperty<ClassificationCategoryType> (commandActivity,
          nameof(IDataProtectionAware.ClassificationCategory),
          parameters.ClassificationCategory,
          (f, v) => f.ToDataProtectionAware().ClassificationCategory = v);
      SetProperty<PublicDisclosureClassificationType> (commandActivity,
          nameof(IPublicDisclosureAware.PublicDisclosureIntention),
          parameters.PublicDisclosureIntention,
          (f, v) => f.ToPublicDisclosureAware().PublicDisclosureIntention = v);
      SetProperty<string> (commandActivity,
          nameof(IPublicDisclosureAware.PublicDisclosureReason),
          parameters.PublicDisclosureReason,
          (f, v) => f.ToPublicDisclosureAware().PublicDisclosureReason = v);
      SetProperty<DeletionModeClassificationType> (commandActivity,
          nameof(IFileHandlingContainerDeletionMixin.DeletionMode),
          parameters.DeletionMode,
          (f, v) => f.ToFileHandlingContainerDeletionMixin().DeletionMode = v);
      SetProperty<bool> (commandActivity, nameof(File.IsPaperFile), parameters.IsPaperFile, (f, v) => f.IsPaperFile = v);
      SetProperty<string> (commandActivity, nameof(File.PaperFileFilingLocation), parameters.PaperFileFilingLocation, (f, v) => f.PaperFileFilingLocation = v);

      var fileType = parameters.FileType?.Invoke<FileTypeClassificationType> (commandActivity, commandActivity.WorkItem);
      if (fileType != null)
      {
        if (file.SubjectArea.CalculatedSubjectAreaFileTypeLinks.All (ftl => ftl.FileType.ID != fileType.ID))
        {
          throw new ActivityCommandException (
                  $"FileType '{fileType.DisplayName}' is not available at subject area '{file.SubjectArea.DisplayName}' in activity '{commandActivity.ID}'.")
              .WithUserMessage (Localization.FileTypeNotAvailableForSubjectArea.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName, fileType.DisplayName));
        }

        AssertIsNotReadOnly (commandActivity, file, nameof(File.FileType));
        file.FileType = fileType;
      }

      var maxNestingLevel = parameters.MaxNestingLevel?.Invoke<int> (commandActivity, commandActivity.WorkItem);
      if (maxNestingLevel != null)
      {
        AssertIsRootFile (commandActivity, file, nameof(parameters.MaxNestingLevel));
        AssertIsNotReadOnly (commandActivity, file, nameof(File.MaxNestingLevel));
        file.MaxNestingLevel = maxNestingLevel;
      }

      var retentionPeriodCategory =
          parameters.RetentionPeriodCategory?.Invoke<RetentionPeriodCategoryClassificationType> (commandActivity, commandActivity.WorkItem);
      if (retentionPeriodCategory != null)
      {
        AssertIsRootFile (commandActivity, file, nameof(parameters.RetentionPeriodCategory));
        file.ToFileDeletionMixin().RetentionPeriodCategory = retentionPeriodCategory;
      }

      var retentionPeriodYears = parameters.RetentionPeriodYears?.Invoke<int> (commandActivity, commandActivity.WorkItem);
      if (retentionPeriodYears != null)
      {
        AssertIsRootFile (commandActivity, file, nameof(parameters.RetentionPeriodYears));
        file.ToFileDeletionMixin().RetentionPeriodYears = retentionPeriodYears.Value;

        if (!file.ToFileDeletionMixin().IsRetentionPeriodYearsInValidRange())
        {
          throw new ActivityCommandException (
                  $"The RetentionPeriodYears value '{retentionPeriodYears}' is not in the valid range in '{commandActivity.ID}'.")
              .WithUserMessage (Localization.ValueNotInValidRangeError.ToLocalizedName()
                  .FormatWith (commandActivity.DisplayName,
                      nameof(parameters.RetentionPeriodYears),
                      file.ToFileDeletionMixin().RetentionPeriodCategory.RetentionPeriodYearsMin,
                      file.ToFileDeletionMixin().RetentionPeriodCategory.RetentionPeriodYearsMax));
        }
      }

      var archivalValueReason = parameters.ArchivalValueReason?.Invoke<string> (commandActivity, commandActivity.WorkItem);
      if (archivalValueReason != null)
      {
        AssertIsRootFile (commandActivity, file, nameof(parameters.ArchivalValueReason));
        AssertIsNotReadOnly (commandActivity, file, nameof(IFileDeletionMixin.ArchivalValueReason));
        file.ToFileDeletionMixin().ArchivalValueReason = archivalValueReason;
      }

      var deletionDelayYears = parameters.DeletionDelayYears?.Invoke<int> (commandActivity, commandActivity.WorkItem);
      if (deletionDelayYears != null)
      {
        AssertIsRootFile (commandActivity, file, nameof(parameters.DeletionDelayYears));
        AssertIsNotReadOnly (commandActivity, file, nameof(IFileDeletionMixin.DeletionDelayYears));
        file.ToFileDeletionMixin().DeletionDelayYears = deletionDelayYears;
      }

      var deletionBackupYears = parameters.DeletionBackupYears?.Invoke<int> (commandActivity, commandActivity.WorkItem);
      if (deletionBackupYears != null)
      {
        AssertIsRootFile (commandActivity, file, nameof(parameters.DeletionBackupYears));
        AssertIsNotReadOnly (commandActivity, file, nameof(IFileDeletionMixin.DeletionBackupYears));
        file.ToFileDeletionMixin().DeletionBackupYears = deletionBackupYears;
      }

      return true;
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        if (commandActivity.EffectiveWorkItemType != typeof (File))
        {
          throw new ActivityCommandException ($"WorkItem type is invalid in activity '{commandActivity.ID}'.")
              .WithUserMessage (Localization.InvalidWorkItemType.ToLocalizedName().FormatWith (commandActivity.DisplayName));
        }

        var parameters = ActivityCommandParser.Parse<SetFileParameters> (commandActivity);
        ShowVersion (commandActivity, parameters);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);

        ValidateBasedataCatalogValues (commandActivity);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    private void ValidateBasedataCatalogValues (CommandActivity commandActivity)
    {
      var parameters = ActivityCommandParser.Parse<SetFileParameters> (commandActivity);

      var fileType = parameters.FileType?.Invoke<FileTypeClassificationType> (commandActivity, commandActivity.WorkItem);
      if (fileType != null)
        AssertIsValidAndActive<FileTypeClassificationType> (commandActivity, fileType, nameof(parameters.FileType));

      try
      {
        parameters.OpeningDate?.Invoke<DateTime> (commandActivity, commandActivity.WorkItem);
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException (
                $"The Parameter {nameof(parameters.OpeningDate)} could not be parsed in '{commandActivity.ID}'. {ex.Message}")
            .WithUserMessage (Localization.DateParsingError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.OpeningDate)));
      }

      try
      {
        parameters.ExternalDate?.Invoke<DateTime> (commandActivity, commandActivity.WorkItem);
      }
      catch (Exception ex)
      {
        throw new ActivityCommandException (
                $"The Parameter {nameof(parameters.ExternalDate)} could not be parsed in '{commandActivity.ID}'. {ex.Message}")
            .WithUserMessage (Localization.DateParsingError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, nameof(parameters.ExternalDate)));
      }

      var leadingGroup = parameters.LeadingGroup?.Invoke<TenantGroup> (commandActivity, commandActivity.WorkItem);
      if (leadingGroup != null)
        AssertIsValidAndActive (commandActivity, leadingGroup, nameof(parameters.LeadingGroup));

      var leadingUser = parameters.LeadingUser?.Invoke<TenantUser> (commandActivity, commandActivity.WorkItem);
      if (leadingUser != null)
        AssertIsValidAndActive (commandActivity, leadingUser, nameof(parameters.LeadingUser));

      var maxNestingLevel = parameters.MaxNestingLevel?.Invoke<int> (commandActivity, commandActivity.WorkItem);
      if (maxNestingLevel < 1)
      {
        throw new ActivityCommandException ($"The MaxNestingLevel must be greater than 1 in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.MaxNestingLevelLessThanOneError.ToLocalizedName().FormatWith (commandActivity.DisplayName));
      }

      var dataProtection = parameters.DataProtection?.Invoke<DataProtectionClassificationType> (commandActivity, commandActivity.WorkItem);
      if (dataProtection != null)
        AssertIsValidAndActive (commandActivity, dataProtection, nameof(parameters.DataProtection));

      var publicDisclosureIntention =
          parameters.PublicDisclosureIntention?.Invoke<PublicDisclosureClassificationType> (commandActivity, commandActivity.WorkItem);
      if (publicDisclosureIntention != null)
        AssertIsValidAndActive (commandActivity, publicDisclosureIntention, nameof(parameters.PublicDisclosureIntention));

      var retentionPeriodCategory =
          parameters.RetentionPeriodCategory?.Invoke<RetentionPeriodCategoryClassificationType> (commandActivity, commandActivity.WorkItem);
      if (retentionPeriodCategory != null)
        AssertIsValidAndActive (commandActivity, retentionPeriodCategory, nameof(parameters.RetentionPeriodCategory));

      var deletionMode = parameters.DeletionMode?.Invoke<DeletionModeClassificationType> (commandActivity, commandActivity.WorkItem);
      if (deletionMode != null)
        AssertIsValidAndActive (commandActivity, deletionMode, nameof(parameters.DeletionMode));

      var deletionDelayYears = parameters.DeletionDelayYears?.Invoke<int> (commandActivity, commandActivity.WorkItem);
      if (deletionDelayYears != null)
      {
        var validRangeDelay = new DeletionDelayYearsRange { DeletionDelayYears = (double)deletionDelayYears };
        new ParameterValidator (commandActivity).ValidateParameters (validRangeDelay);
      }

      var deletionBackupYears = parameters.DeletionBackupYears?.Invoke<int> (commandActivity, commandActivity.WorkItem);
      if (deletionBackupYears != null)
      {
        var validRangeBackup = new DeletionBackupYearsRange { DeletionBackupYears = (double)deletionBackupYears };
        new ParameterValidator (commandActivity).ValidateParameters (validRangeBackup);
      }
    }

    public void AssertIsNotReadOnly (CommandActivity commandActivity, File file, string propertyIdentifier)
    {
      if (file.GetPropertyState (propertyIdentifier).IsReadOnly == true)
      {
        throw new ActivityCommandException (
                $"The property '{propertyIdentifier}' is ReadOnly in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.PropertyIsReadOnlyError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, propertyIdentifier));
      }
    }

    public void AssertIsRootFile (CommandActivity commandActivity, File file, string propertyIdentifier)
    {
      if (!file.IsRootFile)
      {
        throw new ActivityCommandException (
                $"The property '{propertyIdentifier}' can only be set on a root file in activity '{commandActivity.ID}'.")
            .WithUserMessage (Localization.FileIsNotARootFileError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, propertyIdentifier));
      }
    }

    public void AssertIsValidAndActive<T> (CommandActivity commandActivity, T value, string propertyName)
        where T : IActivatableObjectWithTimespan
    {
      if (!value.IsValidAndActive())
      {
        throw new ActivityCommandException (
                $"The value '{value}' for property '{propertyName}' is not active or no longer valid in '{commandActivity.ID}'.")
            .WithUserMessage (Localization.ValueInactiveOrInvalidError.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName, value, propertyName));
      }
    }

    public void SetProperty<T> (CommandActivity commandActivity, string propertyIdentifier, TypedParsedExpression valueExpression, Action<File, T> setter)
    {
      if (!(commandActivity.WorkItem is File file) || valueExpression == null)
        return;

      var value = valueExpression.Invoke<T> (commandActivity, commandActivity.WorkItem);
      if (value != null)
      {
        AssertIsNotReadOnly (commandActivity, file, propertyIdentifier);
        setter (file, value);
      }
    }

    private class DeletionDelayYearsRange
    {
      [Range (-200, 500)]
      public double? DeletionDelayYears { get; set; }
    }

    private class DeletionBackupYearsRange
    {
      [Range (-200, 500)]
      public double? DeletionBackupYears { get; set; }
    }

    private void ShowVersion (CommandActivity commandActivity, SetFileParameters parameters)
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