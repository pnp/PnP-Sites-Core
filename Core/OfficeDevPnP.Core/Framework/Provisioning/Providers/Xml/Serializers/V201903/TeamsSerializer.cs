using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Teams settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201903,
        SerializationSequence = -1, DeserializationSequence = -1,
        Default = false)]
    internal class TeamsSerializer : PnPBaseSchemaSerializer<ProvisioningTeams>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var teams = persistence.GetPublicInstancePropertyValue("Teams");

            if (teams != null)
            {
                var expressions = new Dictionary<Expression<Func<ProvisioningTeams, Object>>, IResolver>();

                // Manage Team Templates
                expressions.Add(t => t.TeamTemplates, new TeamTemplatesFromSchemaToModelTypeResolver());
                expressions.Add(t => t.TeamTemplates[0].JsonTemplate, new ExpressionValueResolver((s, v) => {
                    // Concatenate all the string values in the Text array of strings and return as the content of the JSON template
                    return ((s.GetPublicInstancePropertyValue("Text") as String[])?.Aggregate(String.Empty, (acc, next) => acc += (next != null ? next : String.Empty)));
                }));

                // Manage Teams
                expressions.Add(t => t.Teams, new TeamsFromSchemaToModelTypeResolver());
                expressions.Add(t => t.Teams[0].FunSettings, 
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamFunSettings>("FunSettings"));
                expressions.Add(t => t.Teams[0].GuestSettings, 
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamGuestSettings>("GuestSettings"));
                expressions.Add(t => t.Teams[0].MemberSettings,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamMemberSettings>("MembersSettings"));
                expressions.Add(t => t.Teams[0].MessagingSettings,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamMessagingSettings>("MessagingSettings"));
                expressions.Add(t => t.Teams[0].Security, new TeamSecurityFromSchemaToModelTypeResolver());

                expressions.Add(t => t.Teams[0].Channels[0].Tabs[0].Configuration,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamTabConfiguration>("Configuration"));

                // Handle the JSON content of the Message to send to the channel
                expressions.Add(t => t.Teams[0].Channels[0].Messages[0].Message, new ExpressionValueResolver((s, v) => s));

                PnPObjectsMapper.MapProperties(teams, template.ParentHierarchy.Teams, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            //if (template.Tenant != null && 
            //    (template.Tenant.AppCatalog != null || template.Tenant.ContentDeliveryNetwork != null ||
            //    template.Tenant.SiteDesigns != null || template.Tenant.SiteScripts != null ||
            //    template.Tenant.StorageEntities != null || template.Tenant.Themes != null ||
            //    template.Tenant.WebApiPermissions != null))
            //{
            //    var tenantTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Tenant, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            //    var tenantType = Type.GetType(tenantTypeName, false);
            //    var siteDesignsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteDesignsSiteDesign, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            //    var siteDesignsType = Type.GetType(siteDesignsTypeName, false);
            //    var themeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ThemesTheme, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            //    var themeType = Type.GetType(themeTypeName, false);

            //    if (tenantType != null)
            //    {
            //        var target = Activator.CreateInstance(tenantType, true);

            //        var resolvers = new Dictionary<String, IResolver>();

            //        resolvers.Add($"{tenantType}.AppCatalog",
            //            new AppCatalogFromModelToSchemaTypeResolver());
            //        resolvers.Add($"{tenantType}.ContentDeliveryNetwork",
            //            new CdnFromModelToSchemaTypeResolver());
            //        resolvers.Add($"{siteDesignsType}.SiteScripts",
            //            new SiteScriptRefFromModelToSchemaTypeResolver());

            //        if (themeType != null)
            //        {
            //            resolvers.Add($"{themeType}.Text",
            //                new ExpressionValueResolver((s, v) => {
            //                    return (new String[] { (String)s.GetPublicInstancePropertyValue("Palette") });
            //                }));
            //        }


            //        PnPObjectsMapper.MapProperties(template.Tenant, target, resolvers, recursive: true);

            //        if (target != null &&
            //            (target.GetPublicInstancePropertyValue("AppCatalog") != null ||
            //            target.GetPublicInstancePropertyValue("ContentDeliveryNetwork") != null ||
            //            target.GetPublicInstancePropertyValue("SiteScripts") != null ||
            //            target.GetPublicInstancePropertyValue("SiteDesigns") != null ||
            //            target.GetPublicInstancePropertyValue("StorageEntities") != null ||
            //            target.GetPublicInstancePropertyValue("Themes") != null ||
            //            target.GetPublicInstancePropertyValue("WebApiPermissions") != null))
            //        {
            //            persistence.GetPublicInstanceProperty("Tenant").SetValue(persistence, target);
            //        }
            //    }
            //}
        }
    }
}
