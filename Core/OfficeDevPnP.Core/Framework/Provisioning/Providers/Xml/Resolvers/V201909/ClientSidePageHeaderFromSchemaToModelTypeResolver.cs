using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201909
{
    internal class ClientSidePageHeaderFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public ClientSidePageHeaderFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new ClientSidePageHeader();
            var header = source.GetPublicInstancePropertyValue("Header");

            if (null != header)
            {
                PnPObjectsMapper.MapProperties(header, result, resolvers, recursive);
            }

            return (result);
        }
    }
}
