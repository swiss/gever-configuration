using System;
using System.Reflection;
using ActaNova.Domain;
using JetBrains.Annotations;
using Remotion.Data.DomainObjects;
using Remotion.Security;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Utilities.Autofac;
using Rubicon.Utilities.Expr;
using Rubicon.Utilities.Security;
using Rubicon.Utilities.SecurityManager;

namespace Rubicon.Gever.Bund.Domain.Configs.ActaNova.Domain
{
  [UsedImplicitly]
  public class AccessExpressionAccessModule : ExpressionAccessModule
  {
    protected override void Load (CachingFactory<TypeAspect, Func<MemberInfo, TypeAccessMode>, GetTypeAccess>.Registrator registrator)
    {
      registrator.RegisterTypeAccess(TypeAccessMode.Full, typeof(AccessExpressionExtensions));
    }
  }

  public static class AccessExpressionExtensions
  {
    public static bool HasReadAccess (Document document)
    {
      var user = UserHelper.Current.AsTenantUser();

      if (user.ActualUser == null)
        return false;

      using (new SecurityImpersonationSection(user.ActualUser.GetHandle()))
      {
        var hasAccess = AccessControlUtility.HasAccess(document, GeneralAccessTypes.Read);
        return hasAccess;
      }
    }
  }
}