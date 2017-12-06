using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the ALM settings for a Site Collection
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201801,
        SerializationSequence = 2400, DeserializationSequence = 2400,
        Default = true)]
    internal class ApplicationLifecycleManagementSerializer : PnPBaseSchemaSerializer<ApplicationLifecycleManagement>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var almSettings = persistence.GetPublicInstancePropertyValue("ApplicationLifecycleManagement");

            if (almSettings != null)
            {
                var expressions = new Dictionary<Expression<Func<ApplicationLifecycleManagement, Object>>, IResolver>();

                // Manage the AppCatalog
                expressions.Add(a => a.AppCatalog, new AppCatalogFromSchemaToModelTypeResolver());

                PnPObjectsMapper.MapProperties(almSettings, template.ApplicationLifecycleManagement,
                    expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ApplicationLifecycleManagement != null &&
                (template.ApplicationLifecycleManagement.AppCatalog != null || 
                (template.ApplicationLifecycleManagement.Apps != null &&
                template.ApplicationLifecycleManagement.Apps.Count > 0)))
            {
                var almTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ApplicationLifecycleManagement, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var almType = Type.GetType(almTypeName, false);

                if (almType != null)
                {
                    var target = Activator.CreateInstance(almType, true);

                    var resolvers = new Dictionary<String, IResolver>();

                    resolvers.Add($"{almType}.AppCatalog",
                        new AppCatalogFromModelToSchemaTypeResolver());

                    PnPObjectsMapper.MapProperties(template.ApplicationLifecycleManagement, target,
                        resolvers, recursive: true);

                    if (target != null &&
                        (target.GetPublicInstancePropertyValue("AppCatalog") != null ||
                        target.GetPublicInstancePropertyValue("Apps") != null))
                    {
                        persistence.GetPublicInstanceProperty("ApplicationLifecycleManagement")
                            .SetValue(persistence, target);
                    }
                }
            }
        }
    }
}
