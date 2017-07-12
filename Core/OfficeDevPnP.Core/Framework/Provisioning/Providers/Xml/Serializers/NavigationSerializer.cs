using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Navigation settings
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 800, DeserializationSequence = 800,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class NavigationSerializer : PnPBaseSchemaSerializer<Model.Navigation>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var navigation = persistence.GetPublicInstancePropertyValue("Navigation");

            if (navigation != null)
            {
                var expressions = new Dictionary<Expression<Func<Model.Navigation, Object>>, IResolver>();
                expressions.Add(n => n.CurrentNavigation, new NavigationFromSchemaToModelTypeResolver("CurrentNavigation"));
                expressions.Add(n => n.GlobalNavigation, new NavigationFromSchemaToModelTypeResolver("GlobalNavigation"));

                template.Navigation = new Model.Navigation();
                PnPObjectsMapper.MapProperties(navigation, template.Navigation, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Navigation != null)
            {
                var navigationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Navigation, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var navigationType = Type.GetType(navigationTypeName, true);
                var target = Activator.CreateInstance(navigationType, true);

                var resolvers = new Dictionary<String, IResolver>();

                resolvers.Add($"{navigationType}.GlobalNavigation",
                    new NavigationFromModelToSchemaTypeResolver("GlobalNavigation"));
                resolvers.Add($"{navigationType}.CurrentNavigation",
                    new NavigationFromModelToSchemaTypeResolver("CurrentNavigation"));

                PnPObjectsMapper.MapProperties(template.Navigation, target, resolvers, recursive: true);

                persistence.GetPublicInstanceProperty("Navigation").SetValue(persistence, target);
            }
        }
    }
}
