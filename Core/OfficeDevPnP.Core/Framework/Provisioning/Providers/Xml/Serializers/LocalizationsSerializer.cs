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
    /// Class to serialize/deserialize the Localization Settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        SerializationSequence = -1, DeserializationSequence = -1,
        Default = false)]
    internal class LocalizationsSerializer : PnPBaseSchemaSerializer<Localization>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var localizations = persistence.GetPublicInstancePropertyValue("Localizations");

            if (localizations != null)
            {
                template.Localizations.AddRange(
                    PnPObjectsMapper.MapObjects(localizations,
                            new CollectionFromSchemaToModelTypeResolver(typeof(Localization)))
                            as IEnumerable<Localization>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Localizations != null && template.Localizations.Count > 0)
            {
                var localizationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.LocalizationsLocalization, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var localizationType = Type.GetType(localizationTypeName, true);

                persistence.GetPublicInstanceProperty("Localizations")
                    .SetValue(
                        persistence,
                        PnPObjectsMapper.MapObjects(template.Localizations,
                            new CollectionFromModelToSchemaTypeResolver(localizationType)));
            }
        }
    }
}
