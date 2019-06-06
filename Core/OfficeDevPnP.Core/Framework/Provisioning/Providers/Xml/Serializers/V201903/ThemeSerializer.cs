using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using OfficeDevPnP.Core.Extensions;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Site Footer
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 830, DeserializationSequence = 830,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201903,
        Scope = SerializerScope.ProvisioningTemplate)]
    internal class ThemeSerializer : PnPBaseSchemaSerializer<Theme>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var theme = persistence.GetPublicInstancePropertyValue("Theme");

            if (theme != null)
            {
                var expressions = new Dictionary<Expression<Func<Theme, Object>>, IResolver>();

                // Manage Palette of Theme
                expressions.Add(t => t.Palette, new ExpressionValueResolver((s, v) => {

                    String result = null;

                    if (s != null)
                    {
                        String[] text = s.GetPublicInstancePropertyValue("Text") as String[];
                        if (text != null && text.Length > 0)
                        {
                            result = text.Aggregate(String.Empty, (acc, next) => acc += (next != null ? next : String.Empty));
                        }
                    }

                    return (result?.Trim());
                }));

                template.Theme = new Theme();
                PnPObjectsMapper.MapProperties(theme, template.Theme, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Theme != null)
            {
                var themeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Theme, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var themeType = Type.GetType(themeTypeName, true);
                var target = Activator.CreateInstance(themeType, true);

                var resolvers = new Dictionary<String, IResolver>();

                resolvers.Add($"{themeType}.Text",
                    new ExpressionValueResolver((s, v) => {
                        return (new String[] { (String)s.GetPublicInstancePropertyValue("Palette") });
                    }));

                PnPObjectsMapper.MapProperties(template.Theme, target, resolvers, recursive: true);

                persistence.GetPublicInstanceProperty("Theme").SetValue(persistence, target);
            }
        }
    }
}
