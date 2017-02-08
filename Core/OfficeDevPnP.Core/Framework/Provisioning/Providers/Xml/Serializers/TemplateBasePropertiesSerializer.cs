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
    /// Class to serialize/deserialize the base properties of a template
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 100, DeserializationSequence = 100,
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate), typeof(Xml.V201512.ProvisioningTemplate) },
        AutoInclude = true)]
    internal class TemplateBasePropertiesSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            // TODO: Find a way to avoid using explicit type names

            //// Without lambda expressions
            //var resolvers = new Dictionary<string, IResolver>();
            //resolvers.Add("Version", new FromDecimalToDoubleValueResolver());
            //resolvers.Add("Parameters", new FromArrayToDictionaryValueResolver());
            //PnPObjectsMapper.MapProperties(persistenceTemplate, template, resolvers);

            // Define the dynamic type for the template's properties
            var propertiesTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var propertiesType = Type.GetType(propertiesTypeName, true);
            var propertiesKeySelector = CreateSelectorLambda(propertiesType, "Key");
            var propertiesValueSelector = CreateSelectorLambda(propertiesType, "Value");

            // TODO: Find a way to avoid creating the dictionary of IResolver objects manually
            var expressions = new Dictionary<Expression<Func<ProvisioningTemplate, Object>>, IResolver>();
            expressions.Add(t => t.Version, new FromDecimalToDoubleValueResolver());
            expressions.Add(t => t.Properties,
                new FromArrayToDictionaryValueResolver<String, String>(
                    propertiesType, propertiesKeySelector, propertiesValueSelector));
            PnPObjectsMapper.MapProperties(persistence, template, expressions);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var propertiesTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var propertiesType = Type.GetType(propertiesTypeName, true);

            var keySelector = CreateSelectorLambda(propertiesType, "Key");
            var valueSelector = CreateSelectorLambda(propertiesType, "Value");

            var expressions = new Dictionary<Expression<Func<ProvisioningTemplate, Object>>, IResolver>();
            expressions.Add(t => t.Version, new FromDoubleToDecimalValueResolver());
            expressions.Add(t => t.Properties,
                new FromDictionaryToArrayValueResolver<String, String>(
                    propertiesType, keySelector, valueSelector));
            PnPObjectsMapper.MapProperties(template, persistence, expressions);
        }
    }
}
