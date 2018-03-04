using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Workflows
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 1800, DeserializationSequence = 1800,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class WorkflowsSerializer : PnPBaseSchemaSerializer<Workflows>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var workflows = persistence.GetPublicInstancePropertyValue("Workflows");

            if (workflows != null)
            {
                template.Workflows = new Workflows();
                var expressions = new Dictionary<Expression<Func<Workflows, Object>>, IResolver>();

                expressions.Add(w => w.WorkflowDefinitions[0].Id, new FromStringToGuidValueResolver());
                expressions.Add(w => w.WorkflowDefinitions[0].FormField, new ExpressionValueResolver((s, v) => v != null ? ((XmlElement)v).OuterXml : null));
                expressions.Add(w => w.WorkflowSubscriptions[0].DefinitionId, new FromStringToGuidValueResolver());
                var dictionaryItemType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");
                expressions.Add(w => w.WorkflowDefinitions[0].Properties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                expressions.Add(w => w.WorkflowSubscriptions[0].EventTypes, new ExpressionValueResolver((s, v) =>
                {
                    return (new string[] {
                    (bool)s.GetPublicInstancePropertyValue("ItemAddedEvent") ? "ItemAdded" : null,
                    (bool)s.GetPublicInstancePropertyValue("ItemUpdatedEvent") ? "ItemUpdated" : null,
                    (bool)s.GetPublicInstancePropertyValue("WorkflowStartEvent") ? "WorkflowStart" : null }).Where(e => e != null).ToList();
                }));
                expressions.Add(w => w.WorkflowSubscriptions[0].PropertyDefinitions, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                PnPObjectsMapper.MapProperties(workflows, template.Workflows, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Workflows != null)
            {
                var workflowsType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Workflows, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var workflowDefinitionType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.WorkflowsWorkflowDefinition, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var workflowSubscriptionType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.WorkflowsWorkflowSubscription, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var restrictToType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.WorkflowsWorkflowDefinitionRestrictToType, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                //WorkflowsWorkflowDefinitionRestrictToType.List, wd.RestrictToType
                var target = Activator.CreateInstance(workflowsType, true);

                var expressions = new Dictionary<string, IResolver>();
                expressions.Add($"{workflowDefinitionType}.FormField", new ExpressionValueResolver<string>((v) => v?.ToXmlElement()));
                expressions.Add($"{workflowDefinitionType}.PublishedSpecified", new ExpressionValueResolver<WorkflowDefinition>((s, v) => s.Published));
                expressions.Add($"{workflowDefinitionType}.RequiresAssociationFormSpecified", new ExpressionValueResolver<WorkflowDefinition>((s, v) => s.RequiresAssociationForm));
                expressions.Add($"{workflowDefinitionType}.RequiresInitiationFormSpecified", new ExpressionValueResolver<WorkflowDefinition>((s, v) => s.RequiresInitiationForm));
                expressions.Add($"{workflowDefinitionType}.RestrictToType", new FromStringToEnumValueResolver(restrictToType));
                expressions.Add($"{workflowDefinitionType}.RestrictToTypeSpecified", new ExpressionValueResolver<WorkflowDefinition>((s, v) => !string.IsNullOrEmpty(s.RestrictToType)));
                var dictionaryItemType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");
                expressions.Add($"{workflowDefinitionType}.Properties", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                expressions.Add($"{workflowSubscriptionType}.ManualStartBypassesActivationLimitSpecified", new ExpressionValueResolver<WorkflowSubscription>((s, v) => s.ManualStartBypassesActivationLimit));
                expressions.Add($"{workflowSubscriptionType}.ItemAddedEvent", new ExpressionValueResolver<WorkflowSubscription>((s, v) => s.EventTypes.Contains("ItemAdded")));
                expressions.Add($"{workflowSubscriptionType}.ItemUpdatedEvent", new ExpressionValueResolver<WorkflowSubscription>((s, v) => s.EventTypes.Contains("ItemUpdated")));
                expressions.Add($"{workflowSubscriptionType}.WorkflowStartEvent", new ExpressionValueResolver<WorkflowSubscription>((s, v) => s.EventTypes.Contains("WorkflowStart")));
                expressions.Add($"{workflowSubscriptionType}.PropertyDefinitions", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                PnPObjectsMapper.MapProperties(template.Workflows, target, expressions, recursive: true);

                if (target.GetPublicInstancePropertyValue("WorkflowDefinitions") != null ||
                    target.GetPublicInstancePropertyValue("WorkflowSubscriptions") != null)
                {
                    persistence.GetPublicInstanceProperty("Workflows").SetValue(persistence, target);
                }
            }
        }
    }
}
