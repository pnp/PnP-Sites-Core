using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class SecurityFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            // Use Reflection to get the Security property
            var security = source?.GetType()?.GetProperty("Security",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase)?
                .GetValue(source);

            if (null != security)
            {
                // TODO: Implement
            }

            return (null);
        }
    }
}
