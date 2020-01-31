using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201805;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201903;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201909;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V202002;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers.V202002
{
    /// <summary>
    /// Class to serialize/deserialize the Tenant-wide settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V202002,
        SerializationSequence = 300, DeserializationSequence = 300,
        Scope = SerializerScope.Provisioning)]
    internal class TenantSerializer : PnPBaseSchemaSerializer<ProvisioningTenant>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var tenantSettings = persistence.GetPublicInstancePropertyValue("Tenant");

            if (tenantSettings != null)
            {
                var expressions = new Dictionary<Expression<Func<ProvisioningTenant, Object>>, IResolver>();

                // Manage the AppCatalog
                expressions.Add(t => t.AppCatalog, new AppCatalogFromSchemaToModelTypeResolver());

                // Manage the CDN
                expressions.Add(t => t.ContentDeliveryNetwork, new CdnFromSchemaToModelTypeResolver());

                // Manage the Site Designs mapping with Site Scripts
                expressions.Add(t => t.SiteDesigns[0].SiteScripts, new SiteScriptRefFromSchemaToModelTypeResolver());

                // Manage Palette of Theme
                expressions.Add(t => t.Themes[0].Palette, new ExpressionValueResolver((s, v) => {

                    String result = null;

                    if (s != null)
                    {
                        String[] text = s.GetPublicInstancePropertyValue("Text") as String[];
                        if (text != null && text.Length > 0)
                        {
                            result = text.Aggregate(String.Empty, (acc, next) => acc += (next != null ? next : String.Empty));
                        }
                    }

                    return (result.Trim());
                }));

                // Define the dynamic type for the SP User Profile Properties
                var propertiesTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var propertiesType = Type.GetType(propertiesTypeName, true);
                var propertiesKeySelector = CreateSelectorLambda(propertiesType, "Key");
                var propertiesValueSelector = CreateSelectorLambda(propertiesType, "Value");

                // Manage SP User Profile Properties
                expressions.Add(t => t.SPUsersProfiles[0].Properties,
                    new FromArrayToDictionaryValueResolver<String, String>(
                        propertiesType, propertiesKeySelector, propertiesValueSelector));

                // Manage Office 365 Groups Settings
                expressions.Add(t => t.Office365GroupsSettings,
                    new Office365GroupsSettingsFromSchemaToModel());

                // Manage Sharing Settings
                expressions.Add(t => t.SharingSettings,
                    new SharingSettingsFromSchemaToModel());

                PnPObjectsMapper.MapProperties(tenantSettings, template.Tenant, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Tenant != null && 
                (template.Tenant.AppCatalog != null || template.Tenant.ContentDeliveryNetwork != null ||
                (template.Tenant.SiteDesigns != null && template.Tenant.SiteDesigns.Count > 0) ||
                (template.Tenant.SiteScripts != null && template.Tenant.SiteScripts.Count > 0) ||
                (template.Tenant.StorageEntities != null && template.Tenant.StorageEntities .Count > 0)|| 
                (template.Tenant.Themes != null && template.Tenant.Themes.Count > 0) ||
                (template.Tenant.WebApiPermissions != null && template.Tenant.WebApiPermissions.Count > 0) ||
                template.Tenant.SharingSettings != null))
            {
                var tenantTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Tenant, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var tenantType = Type.GetType(tenantTypeName, false);
                var siteDesignsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteDesignsSiteDesign, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var siteDesignsType = Type.GetType(siteDesignsTypeName, false);
                var themeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Theme, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var themeType = Type.GetType(themeTypeName, false);
                var spUserProfileTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SPUserProfile, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var spUserProfileType = Type.GetType(spUserProfileTypeName, false);

                if (tenantType != null)
                {
                    var target = Activator.CreateInstance(tenantType, true);

                    var resolvers = new Dictionary<String, IResolver>();

                    resolvers.Add($"{tenantType}.AppCatalog",
                        new AppCatalogFromModelToSchemaTypeResolver());
                    resolvers.Add($"{tenantType}.ContentDeliveryNetwork",
                        new CdnFromModelToSchemaTypeResolver());
                    resolvers.Add($"{siteDesignsType}.SiteScripts",
                        new SiteScriptRefFromModelToSchemaTypeResolver());
                    resolvers.Add($"{siteDesignsType}.WebTemplate", 
                        new TenantSiteDesignsWebTemplateFromModelToSchemaValueResolver());

                    if (themeType != null)
                    {
                        resolvers.Add($"{themeType}.Text",
                            new ExpressionValueResolver((s, v) => {
                                return (new String[] { (String)s.GetPublicInstancePropertyValue("Palette") });
                            }));
                    }

                    // Define the dynamic type for the SP User Profile Properties
                    var propertiesTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var propertiesType = Type.GetType(propertiesTypeName, true);

                    var keySelector = CreateSelectorLambda(propertiesType, "Key");
                    var valueSelector = CreateSelectorLambda(propertiesType, "Value");

                    // Manage SP User Profile Properties
                    resolvers.Add($"{spUserProfileType}.Property",
                        new FromDictionaryToArrayValueResolver<String, String>(
                            propertiesType, keySelector, valueSelector, "Properties"));

                    // Manage the Office 365 Groups Settings
                    resolvers.Add($"{tenantType}.Office365GroupsSettings",
                        new Office365GroupsSettingsFromModelToSchema());

                    // Manage the Sharing Settings
                    resolvers.Add($"{tenantType}.SharingSettings",
                        new SharingSettingsFromModelToSchema());

                    PnPObjectsMapper.MapProperties(template.Tenant, target, resolvers, recursive: true);

                    if (target != null &&
                        (target.GetPublicInstancePropertyValue("AppCatalog") != null ||
                        target.GetPublicInstancePropertyValue("ContentDeliveryNetwork") != null ||
                        target.GetPublicInstancePropertyValue("SiteScripts") != null ||
                        target.GetPublicInstancePropertyValue("SiteDesigns") != null ||
                        target.GetPublicInstancePropertyValue("StorageEntities") != null ||
                        target.GetPublicInstancePropertyValue("Themes") != null ||
                        target.GetPublicInstancePropertyValue("WebApiPermissions") != null ||
                        target.GetPublicInstancePropertyValue("SPUsersProfiles") != null ||
                        target.GetPublicInstancePropertyValue("Office365GroupLifecyclePolicies") != null ||
                        target.GetPublicInstancePropertyValue("Office365GroupsSettings") != null ||
                        target.GetPublicInstancePropertyValue("SharingSettings") != null))
                    {
                        persistence.GetPublicInstanceProperty("Tenant").SetValue(persistence, target);
                    }
                }
            }
        }
    }
}
