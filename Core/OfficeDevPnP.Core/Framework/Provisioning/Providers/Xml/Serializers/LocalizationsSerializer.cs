using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the localization settings
    /// </summary>
    [TemplateSchemaSerializer(
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate), typeof(Xml.V201512.ProvisioningTemplate) },
        AutoInclude = false)]
    internal class LocalizationsSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var localizations = persistence.GetType().GetProperty("Localizations",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.IgnoreCase |
                System.Reflection.BindingFlags.Public).GetValue(persistence);

            if (localizations != null)
            {
                template.Localizations.AddRange(
                    PnPObjectsMapper.MapObject(localizations,
                            new CollectionFromSchemaToModelTypeResolver(typeof(Localization)))
                            as IEnumerable<Localization>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var localizationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.LocalizationsLocalization, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var localizationType = Type.GetType(localizationTypeName, true);

            persistence.GetType().GetProperty("Localizations",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.IgnoreCase |
                System.Reflection.BindingFlags.Public).SetValue(
                    persistence,
                    PnPObjectsMapper.MapObject(template.Localizations,
                        new CollectionFromModelToSchemaTypeResolver(localizationType)));
        }
    }
}
