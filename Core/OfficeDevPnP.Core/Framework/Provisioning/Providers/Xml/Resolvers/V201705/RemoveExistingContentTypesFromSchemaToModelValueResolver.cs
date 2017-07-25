using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    /// <summary>
    /// Resolves the RemoveExistingContentTypes attribute from Schema to Domain Model
    /// </summary>
    internal class RemoveExistingContentTypesFromSchemaToModelValueResolver : IValueResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            var result = false;

            var allowedContentTypes = source.GetPublicInstancePropertyValue("AllowedContentTypes");
            var removeExistingContentTypes = allowedContentTypes?.GetPublicInstancePropertyValue("RemoveExistingContentTypes");

            if (null != removeExistingContentTypes)
            {
                result = (Boolean)removeExistingContentTypes;
            }

            return (result);
        }
    }
}
