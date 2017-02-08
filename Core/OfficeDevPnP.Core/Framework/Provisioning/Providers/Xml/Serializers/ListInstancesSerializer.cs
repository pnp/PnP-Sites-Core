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
    /// Class to serialize/deserialize the list instances
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 300, DeserializationSequence = 300,
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate), typeof(Xml.V201512.ProvisioningTemplate) },
        Default = true)]
    internal class ListInstancesSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var lists = persistence.GetType().GetProperty("Lists",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.IgnoreCase |
                System.Reflection.BindingFlags.Public).GetValue(persistence);

            template.Lists.AddRange(
                PnPObjectsMapper.MapObject(lists,
                        new CollectionFromSchemaToModelTypeResolver(typeof(ListInstance)), recursive: true)
                        as IEnumerable<ListInstance>);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var listInstanceTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstance, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var listInstanceType = Type.GetType(listInstanceTypeName, true);

            persistence.GetType().GetProperty("Lists",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.IgnoreCase |
                System.Reflection.BindingFlags.Public).SetValue(
                    persistence,
                    PnPObjectsMapper.MapObject(template.Lists,
                        new CollectionFromModelToSchemaTypeResolver(listInstanceType), recursive: true));
        }
    }
}
