using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers.V201903
{
    /// <summary>
    /// Class to serialize/deserialize the Security settings
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 700, DeserializationSequence = 700,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201903,
        Scope = SerializerScope.ProvisioningTemplate)]
    internal class SecuritySerializer : PnPBaseSchemaSerializer<SiteSecurity>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var security = persistence.GetPublicInstancePropertyValue("Security");

            if (security != null)
            {
                var expressions = new Dictionary<Expression<Func<SiteSecurity, Object>>, IResolver>();

                expressions.Add(s => s.SiteSecurityPermissions, new PropertyObjectTypeResolver<SiteSecurity>(s => s.SiteSecurityPermissions, o => o.GetPublicInstancePropertyValue("Permissions")));
                expressions.Add(s => s.SiteSecurityPermissions.RoleDefinitions[0].Permissions, 
                    new ExpressionCollectionValueResolver<PermissionKind>((i) => (PermissionKind)Enum.Parse(typeof(PermissionKind), i.ToString())));

                expressions.Add(s => s.AdditionalAdministrators,
                    new TemplateSecurityUsersFromSchemaToModelTypeResolver("AdditionalAdministrators"));
                expressions.Add(s => s.ClearExistingAdministrators,
                    new ExpressionValueResolver((s, p) =>
                    {
                        var clearValues = s?.GetPublicInstancePropertyValue("AdditionalAdministrators")?.GetPublicInstancePropertyValue("ClearExistingItems");
                        return(clearValues != null ? (Boolean)clearValues : false);
                    }));
                expressions.Add(s => s.AdditionalOwners,
                    new TemplateSecurityUsersFromSchemaToModelTypeResolver("AdditionalOwners"));
                expressions.Add(s => s.ClearExistingOwners,
                    new ExpressionValueResolver((s, p) =>
                    {
                        var clearValues = s?.GetPublicInstancePropertyValue("AdditionalOwners")?.GetPublicInstancePropertyValue("ClearExistingItems");
                        return (clearValues != null ? (Boolean)clearValues : false);
                    }));
                expressions.Add(s => s.AdditionalMembers,
                    new TemplateSecurityUsersFromSchemaToModelTypeResolver("AdditionalMembers"));
                expressions.Add(s => s.ClearExistingMembers,
                    new ExpressionValueResolver((s, p) =>
                    {
                        var clearValues = s?.GetPublicInstancePropertyValue("AdditionalMembers")?.GetPublicInstancePropertyValue("ClearExistingItems");
                        return (clearValues != null ? (Boolean)clearValues : false);
                    }));
                expressions.Add(s => s.AdditionalVisitors,
                    new TemplateSecurityUsersFromSchemaToModelTypeResolver("AdditionalVisitors"));
                expressions.Add(s => s.ClearExistingVisitors,
                    new ExpressionValueResolver((s, p) =>
                    {
                        var clearValues = s?.GetPublicInstancePropertyValue("AdditionalVisitors")?.GetPublicInstancePropertyValue("ClearExistingItems");
                        return (clearValues != null ? (Boolean)clearValues : false);
                    }));
                expressions.Add(s => s.SiteGroups[0].Members,
                    new TemplateSecurityUsersFromSchemaToModelTypeResolver("Members"));
                expressions.Add(s => s.SiteGroups[0].ClearExistingMembers,
                    new ExpressionValueResolver((s, p) =>
                    {
                        var clearValues = s?.GetPublicInstancePropertyValue("Members")?.GetPublicInstancePropertyValue("ClearExistingItems");
                        return (clearValues != null ? (Boolean)clearValues : false);
                    }));

                PnPObjectsMapper.MapProperties(security, template.Security, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Security != null)
            {
                var securityTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Security, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var securityType = Type.GetType(securityTypeName, true);

                var expressions = new Dictionary<string, IResolver>();

                expressions.Add($"{securityType}.BreakRoleInheritanceSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{securityType}.ResetRoleInheritanceSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{securityType}.CopyRoleAssignmentsSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{securityType}.RemoveExistingUniqueRoleAssignmentsSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{securityType}.ClearSubscopesSpecified", new ExpressionValueResolver(() => true));

                var securityPermissionsType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SecurityPermissions, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}");
                var roleDefinitionType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.RoleDefinition, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}");
                var roleDefinitionPermissionType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.RoleDefinitionPermission, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}");
                var siteGroupType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteGroup, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}");

                var target = Activator.CreateInstance(securityType, true);

                expressions.Add($"{securityType}.Permissions", new PropertyObjectTypeResolver(securityPermissionsType, "SiteSecurityPermissions"));
                expressions.Add($"{roleDefinitionType}.Permissions", new ExpressionCollectionValueResolver((i) => Enum.Parse(roleDefinitionPermissionType, i.ToString()), roleDefinitionPermissionType));

                expressions.Add($"{siteGroupType}.AllowMembersEditMembershipSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{siteGroupType}.AllowRequestToJoinLeaveSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{siteGroupType}.AutoAcceptRequestToJoinLeaveSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{siteGroupType}.OnlyAllowMembersViewMembershipSpecified", new ExpressionValueResolver(() => true));

                expressions.Add($"{securityType}.AdditionalAdministrators",
                    new TemplateSecurityUsersFromModelToSchemaTypeResolver("AdditionalAdministrators", "ClearExistingAdministrators"));
                expressions.Add($"{securityType}.AdditionalOwners",
                    new TemplateSecurityUsersFromModelToSchemaTypeResolver("AdditionalOwners", "ClearExistingOwners"));
                expressions.Add($"{securityType}.AdditionalMembers",
                    new TemplateSecurityUsersFromModelToSchemaTypeResolver("AdditionalMembers", "ClearExistingMembers"));
                expressions.Add($"{securityType}.AdditionalVisitors",
                    new TemplateSecurityUsersFromModelToSchemaTypeResolver("AdditionalVisitors", "ClearExistingVisitors"));
                expressions.Add($"{siteGroupType}.Members",
                    new TemplateSecurityUsersFromModelToSchemaTypeResolver("Members", "ClearExistingMembers"));

                PnPObjectsMapper.MapProperties(template.Security, target, expressions, recursive: true);

                if (target != null &&
                    (target.GetPublicInstancePropertyValue("AdditionalAdministrators") != null ||
                    target.GetPublicInstancePropertyValue("AdditionalMembers") != null ||
                    target.GetPublicInstancePropertyValue("AdditionalOwners") != null ||
                    target.GetPublicInstancePropertyValue("AdditionalVisitors") != null ||
                    target.GetPublicInstancePropertyValue("SiteGroups") != null ||
                    (target.GetPublicInstancePropertyValue("Permissions") != null &&
                    (
                        target.GetPublicInstancePropertyValue("Permissions").GetPublicInstancePropertyValue("RoleDefinitions") != null && (((Array)target.GetPublicInstancePropertyValue("Permissions").GetPublicInstancePropertyValue("RoleDefinitions")).Length > 0) ||
                        target.GetPublicInstancePropertyValue("Permissions").GetPublicInstancePropertyValue("RoleAssignments") != null && (((Array)target.GetPublicInstancePropertyValue("Permissions").GetPublicInstancePropertyValue("RoleAssignments")).Length > 0)
                    ))))
                {
                    persistence.GetPublicInstanceProperty("Security").SetValue(persistence, target);
                }
            }
        }
    }
}
