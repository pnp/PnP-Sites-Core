using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the list instances
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 300, DeserializationSequence = 300,
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate) },
        Default = true)]
    internal class ListInstancesSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var lists = persistence.GetPublicInstancePropertyValue("Lists");

            var expressions = new Dictionary<Expression<Func<ListInstance, Object>>, IResolver>();

            // Define custom resolver for FieldRef.ID because needs conversion from String to GUID
            expressions.Add(l => l.FieldRefs[0].Id, new FromStringToGuidValueResolver());

            // Define custom resolvers for DataRows Values and Security
            var dataRowValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DataValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var dataRowValueType = Type.GetType(dataRowValueTypeName, true);
            var dataRowValueKeySelector = CreateSelectorLambda(dataRowValueType, "FieldName");
            var dataRowValueValueSelector = CreateSelectorLambda(dataRowValueType, "Value");
            expressions.Add(l => l.DataRows[0].Values,
                new FromArrayToDictionaryValueResolver<String, String>(
                    dataRowValueType, dataRowValueKeySelector, dataRowValueValueSelector));

            expressions.Add(l => l.DataRows[0].Security, new SecurityFromSchemaToModelTypeResolver());

            // Define custom resolver for Fields Defaults
            var fieldDefaultTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.FieldDefault, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var fieldDefaultType = Type.GetType(fieldDefaultTypeName, true);
            var fieldDefaultKeySelector = CreateSelectorLambda(fieldDefaultType, "FieldName");
            var fieldDefaultValueSelector = CreateSelectorLambda(fieldDefaultType, "Value");
            expressions.Add(l => l.FieldDefaults,
                new FromArrayToDictionaryValueResolver<String, String>(
                    fieldDefaultType, fieldDefaultKeySelector, fieldDefaultValueSelector));

            // Define custom resolver for Security
            expressions.Add(l => l.Security, new SecurityFromSchemaToModelTypeResolver());

            // TODO: Define custom resolver for UserCustomActions > CommandUIExtension (XML Any)

            // TODO: Define custom resolver for Views (XML Any)

            // TODO: Define custom resolver for recursive Folders

            template.Lists.AddRange(
                PnPObjectsMapper.MapObjects<ListInstance>(lists,
                        new CollectionFromSchemaToModelTypeResolver(typeof(ListInstance)), 
                        expressions, 
                        recursive: true)
                        as IEnumerable<ListInstance>);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var listInstanceTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstance, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var listInstanceType = Type.GetType(listInstanceTypeName, true);

            // TODO: Define the way back for the non-standard properties defined in the Deserialize method

            persistence.GetPublicInstanceProperty("Lists")
                .SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.Lists,
                        new CollectionFromModelToSchemaTypeResolver(listInstanceType), recursive: true));
        }
    }
}
