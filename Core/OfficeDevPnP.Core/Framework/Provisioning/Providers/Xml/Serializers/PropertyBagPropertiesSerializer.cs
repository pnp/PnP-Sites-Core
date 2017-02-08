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
    /// Class to serialize/deserialize the property bag properties
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 200, DeserializationSequence = 200,
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate), typeof(Xml.V201512.ProvisioningTemplate) },
        AutoInclude = true)]
    internal class PropertyBagPropertiesSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var properties = persistence.GetType().GetProperty("PropertyBagEntries",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.IgnoreCase |
                System.Reflection.BindingFlags.Public).GetValue(persistence);

            template.PropertyBagEntries.AddRange(
                PnPObjectsMapper.MapObject(properties,
                        new CollectionFromSchemaToModelTypeResolver(typeof(PropertyBagEntry)))
                        as IEnumerable<PropertyBagEntry>);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var propertyBagTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PropertyBagEntry, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var propertyBagType = Type.GetType(propertyBagTypeName, true);

            persistence.GetType().GetProperty("PropertyBagEntries",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.IgnoreCase |
                System.Reflection.BindingFlags.Public).SetValue(
                    persistence,
                    PnPObjectsMapper.MapObject(template.PropertyBagEntries,
                        new CollectionFromModelToSchemaTypeResolver(propertyBagType)));
        }
    }
}
