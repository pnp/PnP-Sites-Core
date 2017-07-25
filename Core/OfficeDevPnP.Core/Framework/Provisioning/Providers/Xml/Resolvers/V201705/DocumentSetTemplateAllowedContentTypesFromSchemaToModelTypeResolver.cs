using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    /// <summary>
    /// Resolves a list of Allowed Content Types from Schema to Domain Model
    /// </summary>
    internal class DocumentSetTemplateAllowedContentTypesFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public DocumentSetTemplateAllowedContentTypesFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new List<String>();

            var allowedContentTypesContainer = source.GetPublicInstancePropertyValue("AllowedContentTypes");
            var allowedContentTypes = allowedContentTypesContainer?.GetPublicInstancePropertyValue("AllowedContentType");

            if (null != allowedContentTypes)
            {
                foreach(var ac in (IEnumerable)allowedContentTypes)
                {
                    var contentTypeId = ac?.GetPublicInstancePropertyValue("ContentTypeID");
                    result.Add((String)contentTypeId);
                }
            }

            return (result);
        }
    }
}
