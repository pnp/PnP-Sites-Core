using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201805;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Tenant-wide settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201807,
        SerializationSequence = -1, DeserializationSequence = -1,
        Default = false)]
    internal class SequenceSerializer : PnPBaseSchemaSerializer<ProvisioningSequence>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var sequences = persistence.GetPublicInstancePropertyValue("Sequence");

            if (sequences != null)
            {
                var expressions = new Dictionary<Expression<Func<ProvisioningSequence, Object>>, IResolver>();

                expressions.Add(seq => seq.TermStore, new ExpressionValueResolver((s, v) => {

                    var termGroupsExpressions = new Dictionary<Expression<Func<TermGroup, Object>>, IResolver>();
                    termGroupsExpressions.Add(g => g.Id, new FromStringToGuidValueResolver());
                    termGroupsExpressions.Add(g => g.TermSets[0].Id, new FromStringToGuidValueResolver());
                    termGroupsExpressions.Add(g => g.TermSets[0].Terms[0].Id, new FromStringToGuidValueResolver());
                    termGroupsExpressions.Add(g => g.TermSets[0].Terms[0].SourceTermId, new FromStringToGuidValueResolver());

                    var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                    var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                    var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");
                    termGroupsExpressions.Add(g => g.TermSets[0].Properties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));
                    termGroupsExpressions.Add(g => g.TermSets[0].Terms[0].LocalProperties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "LocalCustomProperties"));
                    termGroupsExpressions.Add(g => g.TermSets[0].Terms[0].Properties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "CustomProperties"));
                    termGroupsExpressions.Add(g => g.TermSets[0].Terms[0].Terms,
                        new PropertyObjectTypeResolver<Term>(t => t.Terms,
                        o => o.GetPublicInstancePropertyValue("Terms")?.GetPublicInstancePropertyValue("Items"),
                        new CollectionFromSchemaToModelTypeResolver(typeof(Term))));

                    var result = new Model.ProvisioningTermStore();
                    result.TermGroups.AddRange(
                        PnPObjectsMapper.MapObjects<TermGroup>(v,
                            new CollectionFromSchemaToModelTypeResolver(typeof(TermGroup)),
                            termGroupsExpressions, recursive: true)
                            as IEnumerable<TermGroup>);

                    return (result);
                }));

                template.ParentHierarchy.Sequences.AddRange(
                    PnPObjectsMapper.MapObjects<ProvisioningSequence>(sequences,
                            new CollectionFromSchemaToModelTypeResolver(typeof(ProvisioningSequence)),
                            expressions, recursive: true)
                            as IEnumerable<ProvisioningSequence>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ParentHierarchy != null && 
                template.ParentHierarchy.Sequences != null &&
                template.ParentHierarchy.Sequences.Count > 0)
            {
                var sequenceTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Sequence, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var sequenceType = Type.GetType(sequenceTypeName, true);

                persistence.GetPublicInstanceProperty("Sequence")
                    .SetValue(
                        persistence,
                        PnPObjectsMapper.MapObjects(template.ParentHierarchy.Sequences,
                            new CollectionFromModelToSchemaTypeResolver(sequenceType), recursive: true));
            }
        }
    }
}
