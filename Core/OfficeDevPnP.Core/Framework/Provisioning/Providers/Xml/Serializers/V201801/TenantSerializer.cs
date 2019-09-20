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

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Tenant-wide settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201801,
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

                PnPObjectsMapper.MapProperties(tenantSettings, template.Tenant, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Tenant != null && 
                (template.Tenant.AppCatalog != null || template.Tenant.ContentDeliveryNetwork != null ||
                template.Tenant.SiteDesigns != null || template.Tenant.SiteScripts != null ||
                template.Tenant.StorageEntities != null || template.Tenant.Themes != null ||
                template.Tenant.WebApiPermissions != null))
            {
                var tenantTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Tenant, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var tenantType = Type.GetType(tenantTypeName, false);
                var siteDesignsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteDesignsSiteDesign, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var siteDesignsType = Type.GetType(siteDesignsTypeName, false);
                var themeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ThemesTheme, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var themeType = Type.GetType(themeTypeName, false);

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


                    PnPObjectsMapper.MapProperties(template.Tenant, target, resolvers, recursive: true);

                    if (target != null &&
                        (target.GetPublicInstancePropertyValue("AppCatalog") != null ||
                        target.GetPublicInstancePropertyValue("ContentDeliveryNetwork") != null ||
                        target.GetPublicInstancePropertyValue("SiteScripts") != null ||
                        target.GetPublicInstancePropertyValue("SiteDesigns") != null ||
                        target.GetPublicInstancePropertyValue("StorageEntities") != null ||
                        target.GetPublicInstancePropertyValue("Themes") != null ||
                        target.GetPublicInstancePropertyValue("WebApiPermissions") != null))
                    {
                        persistence.GetPublicInstanceProperty("Tenant").SetValue(persistence, target);
                    }
                }
            }
        }
    }
}
