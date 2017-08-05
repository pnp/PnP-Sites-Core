using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Web Settings
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 300, DeserializationSequence = 300,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class WebSettingsSerializer : PnPBaseSchemaSerializer<WebSettings>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var webSettings = persistence.GetPublicInstancePropertyValue("WebSettings");

            if (webSettings != null)
            {
                template.WebSettings = new WebSettings();
                PnPObjectsMapper.MapProperties(webSettings, template.WebSettings, null, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.WebSettings != null)
            {
                var webSettingsType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.WebSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var target = Activator.CreateInstance(webSettingsType, true);
                var expressions = new Dictionary<string, IResolver>();
                expressions.Add($"{webSettingsType}.NoCrawlSpecified", new ExpressionValueResolver(() => true));

                PnPObjectsMapper.MapProperties(template.WebSettings, target, expressions, recursive: true);

                persistence.GetPublicInstanceProperty("WebSettings").SetValue(persistence, target);
            }
        }
    }
}
