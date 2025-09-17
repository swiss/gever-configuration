// Name: Eingang registrieren (RegisterIncoming)
// Version: 2.0.1
// Datum: 16.09.2025
// Autor: RUBICON IT - Claudia Fleck - claudia.fleck@rubicon.eu; Ursprünglich LHIND - Farid Modarressi - farid.modarressi@lhind.dlh.de - Lufthansa Industry Solutions GmbH

// Mit dieser Befehlsaktivität wird ein Eingangsobjekt gemäss Parameter registriert.
// Version 1.2.1 - farid.modarressi@lhind.dlh.de - 20.11.2023

using ActaNova.Domain;
using ActaNova.Domain.Workflow;
using Remotion.Globalization;
using Rubicon.Domain;
using Rubicon.Gever.Bund.Domain.Workflow;
using Rubicon.Gever.Bund.Domain.Workflow.ValidationTools;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Globalization;
using Rubicon.Workflow.Domain;
using System;
using System.Reflection;

namespace Bukonf.Gever.Workflow
{
  [VersionInfo ("2.0.1", "16.09.2025", "RUBICON IT - Claudia Fleck - claudia.fleck@rubicon.eu - Ursprünglich Farid Modarressi LHIND")]
  public class RegisterIncomingCommandModule : GeverActivityCommandModule
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

    public class RegisterIncomingParameters : GeverActivityCommandParameters
    {
      [RequiredOneOfMany (new[] { "TargetReferenceID", "Target" })]
      [EitherOr (new[] { "TargetReferenceID", "Target" })]
      [ExpectedExpressionType (typeof (string))]
      public TypedParsedExpression TargetReferenceID { get; set; }

      [RequiredOneOfMany (new[] { "TargetReferenceID", "Target" })]
      [EitherOr (new[] { "TargetReferenceID", "Target" })]
      [ExpectedExpressionType (typeof (File))]
      public TypedParsedExpression Target { get; set; }

      [ExpectedExpressionType (typeof (bool))]
      public TypedParsedExpression ShowVersion { get; set; }
    }

    [LocalizationEnum]
    public enum ModuleLocalization
    {
      [De("Befehlsaktivität \"{0}\": Das Zieldossier konnte nicht gefunden werden."),
       MultiLingualName("Activité de commande \"{0}\": Le dossier cible de l'activité n'a pas pu être trouvé.", "Fr"),
       MultiLingualName("Attività di comando \"{0}\": Non è stato possibile trovare il dossier di destinazione dell'attività.", "It")]
      NoFile,

      [De("Befehlsaktivität \"{0}\": Der Geschäftsobjekt der Aktivität \"{1}\" ist bereits registriert."),
       MultiLingualName("Activité de commande \"{0}\": L’objet métier de l’activité \"{1}\" est déjà enregistré.", "Fr"),
       MultiLingualName("Attività di comando \"{0}\": L'oggetto business dell'attività \"{1}\" è già registrato.", "It")]
      IncomingAlreadyRegistered,

      [De("Befehlsaktivität \"{0}\": Das Geschäftsobjekt \"{1}\" der Aktivität kann nicht bearbeitet werden."),
       MultiLingualName("Activité de commande \"{0}\": L’objet métier \"{1}\" de l’activité ne peut être édité.", "Fr"),
       MultiLingualName("Attività di comando \"{0}\": L'oggetto business \"{1}\" dell'attività non può essere modificato.", "It")]
      NotEditable,

      [De("Befehlsaktivität \"{0}\": Version: \"{1}\" - Datum: \"{2}\" - Autor \"{3}\"")]
      ShowVersion,
    }

    public RegisterIncomingCommandModule () : base (
        "RegisterIncoming:ActivityCommandClassificationType")
    {
    }

    public override bool TryExecute (CommandActivity commandActivity)
    {
      var hostObject = (Incoming)commandActivity.WorkItem;
      if (!hostObject.CanBeInserted())
      {
        throw new ActivityCommandException ($"Incoming already registered {hostObject}").WithUserMessage (ModuleLocalization
            .IncomingAlreadyRegistered
            .ToLocalizedName()
            .FormatWith (commandActivity.DisplayName, commandActivity.WorkItem));
      }

      var targetFile = GetTargetFile (commandActivity);

      if (targetFile.CanEdit())
      {
        hostObject.Insert (targetFile);
        return true;
      }

      throw new ActivityCommandException ($"User cannot edit {targetFile}")
          .WithUserMessage (ModuleLocalization.NotEditable
              .ToLocalizedName().FormatWith (commandActivity.DisplayName, targetFile));
    }

    public override string Validate (CommandActivity commandActivity)
    {
      try
      {
        WorkItemTypeValidator typeValidator = new WorkItemTypeValidator();
        typeValidator.AddAllowedTypes (typeof (Incoming));
        typeValidator.IsValidWorkItemType (commandActivity);
        var parameters = ActivityCommandParser.Parse<RegisterIncomingParameters> (commandActivity);
        ShowVersion (commandActivity, parameters);
        new ParameterValidator (commandActivity).ValidateParameters (parameters);
      }
      catch (Exception ex)
      {
        return ex.GetUserMessage() ?? ex.Message;
      }

      return null;
    }

    private File GetTargetFile (CommandActivity commandActivity)
    {
      File file = null;
      var parameters = ActivityCommandParser.Parse<RegisterIncomingParameters> (commandActivity);

      if (parameters?.Target != null)
      {
        file = parameters.Target.Invoke<File> (commandActivity, commandActivity.WorkItem);
      }
      else if (parameters?.TargetReferenceID != null)
      {
        var targetFileRef = parameters.TargetReferenceID.Invoke<string> (commandActivity, commandActivity.WorkItem);
        file = new ReferenceHandle<File> (targetFileRef).TryGetObject (null, true);
      }

      if (file == null)
      {
        throw new ActivityCommandException ($"Could not resolve target file in activity '{commandActivity.ID}'.")
            .WithUserMessage (ModuleLocalization.NoFile.ToLocalizedName()
                .FormatWith (commandActivity.DisplayName));
      }

      return file;
    }

    private void ShowVersion (CommandActivity commandActivity, RegisterIncomingParameters parameters)
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