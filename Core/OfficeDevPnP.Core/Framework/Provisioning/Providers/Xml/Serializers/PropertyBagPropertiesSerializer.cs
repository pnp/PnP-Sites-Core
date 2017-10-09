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
    /// Class to serialize/deserialize the Property Bag Properties
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 200, DeserializationSequence = 200,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class PropertyBagPropertiesSerializer : PnPBaseSchemaSerializer<PropertyBagEntry>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var properties = persistence.GetPublicInstancePropertyValue("PropertyBagEntries");

            if (properties != null)
            {
                template.PropertyBagEntries.AddRange(
                    PnPObjectsMapper.MapObjects(properties,
                            new CollectionFromSchemaToModelTypeResolver(typeof(PropertyBagEntry)))
                            as IEnumerable<PropertyBagEntry>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.PropertyBagEntries != null && template.PropertyBagEntries.Count > 0)
            {
                var propertyBagTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PropertyBagEntry, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var propertyBagType = Type.GetType(propertyBagTypeName, true);


                var expressions = new Dictionary<string, IResolver>();
                expressions.Add($"{propertyBagType}.OverwriteSpecified", new ExpressionValueResolver(() => true));

                persistence.GetPublicInstanceProperty("PropertyBagEntries")
                    .SetValue(
                        persistence,
                        PnPObjectsMapper.MapObjects(template.PropertyBagEntries,
                        new CollectionFromModelToSchemaTypeResolver(propertyBagType),
                        expressions));
            }
        }
    }
}
