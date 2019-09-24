using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers.V201909
{
    /// <summary>
    /// Class to serialize/deserialize the Site Settings
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 350, DeserializationSequence = 350,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201909,
        Scope = SerializerScope.ProvisioningTemplate)]
    internal class SiteSettingsSerializer : PnPBaseSchemaSerializer<SiteSettings>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var webSettings = persistence.GetPublicInstancePropertyValue("SiteSettings");

            if (webSettings != null)
            {
                template.SiteSettings = new SiteSettings();
                PnPObjectsMapper.MapProperties(webSettings, template.SiteSettings, null, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.SiteSettings != null)
            {
                var webSettingsType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var target = Activator.CreateInstance(webSettingsType, true);
                var expressions = new Dictionary<string, IResolver>();

                PnPObjectsMapper.MapProperties(template.SiteSettings, target, expressions, recursive: true);

                persistence.GetPublicInstanceProperty("SiteSettings").SetValue(persistence, target);
            }
        }
    }
}
