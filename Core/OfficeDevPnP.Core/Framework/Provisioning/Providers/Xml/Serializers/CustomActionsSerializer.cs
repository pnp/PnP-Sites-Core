using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the content types
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 400, DeserializationSequence = 400,
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate) },
        Default = true)]
    internal class CustomActionsSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var customActions = persistence.GetPublicInstancePropertyValue("CustomActions");
            var expressions = new Dictionary<Expression<Func<CustomActions, Object>>, IResolver>();

            expressions.Add(c => c.SiteCustomActions[0].CommandUIExtension, new XmlAnyFromSchemaToModelValueResolver("CommandUIExtension"));
            expressions.Add(c => c.SiteCustomActions[0].RegistrationType, new FromStringToEnumValueResolver(typeof(UserCustomActionRegistrationType)));
            expressions.Add(c => c.SiteCustomActions[0].Rights, new FromStringToBasePermissionsValueResolver());

            PnPObjectsMapper.MapProperties(customActions, template.CustomActions, expressions, true);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var customActionsName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CustomActions, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var customActionName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CustomAction, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var customActionsType = Type.GetType(customActionsName, true);

            var target = Activator.CreateInstance(customActionsType, true);

            var expressions = new Dictionary<string, IResolver>();
            //expressions.Add($"{customActionName}.CommandUIExtension", new Comm);

            PnPObjectsMapper.MapProperties(template.CustomActions, target, expressions, recursive: true);

            //expressions.Add(c=> c.SiteCustomActions[0].Sequence, )
            //expressions.Add(c => c.SiteCustomActions[0].CommandUIExtension, new XmlAnyFromSchemaToModelValueResolver("CommandUIExtension"));
            //expressions.Add(c => c.SiteCustomActions[0].RegistrationType, new FromStringToEnumValueResolver(typeof(UserCustomActionRegistrationType)));
            //expressions.Add(c => c.SiteCustomActions[0].Rights, new FromStringToBasePermissionsValueResolver());

            persistence.GetPublicInstanceProperty("CustomActions").SetValue(persistence, target);
        }
    }
}
