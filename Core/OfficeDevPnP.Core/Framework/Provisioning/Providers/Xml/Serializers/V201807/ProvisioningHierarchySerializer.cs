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
        Scope = SerializerScope.Tenant)]
    internal class ProvisioningHierarchySerializer : PnPBaseSchemaSerializer<ProvisioningHierarchy>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            if (persistence != null)
            {
                var expressions = new Dictionary<Expression<Func<ProvisioningHierarchy, Object>>, IResolver>();
                PnPObjectsMapper.MapProperties(persistence, template.ParentHierarchy, expressions, recursive: false);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            //if (template.ParentHierarchy != null)
            //{
            //    var resolvers = new Dictionary<String, IResolver>();
            //    PnPObjectsMapper.MapProperties(template.ParentHierarchy, persistence, resolvers, recursive: false);
            //}
        }
    }
}
