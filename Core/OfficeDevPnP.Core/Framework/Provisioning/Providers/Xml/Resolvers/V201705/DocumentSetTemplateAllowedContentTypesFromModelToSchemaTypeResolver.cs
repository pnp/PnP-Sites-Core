using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class DocumentSetTemplateAllowedContentTypesFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public DocumentSetTemplateAllowedContentTypesFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            var documentSetTemplate = source as Model.DocumentSetTemplate;

            if (null != documentSetTemplate)
            {
                var allowedContentTypesTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DocumentSetTemplateAllowedContentTypes, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var allowedContentTypesType = Type.GetType(allowedContentTypesTypeName, true);
                result = Activator.CreateInstance(allowedContentTypesType);

                var allowedContentTypeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DocumentSetTemplateAllowedContentTypesAllowedContentType, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var allowedContentTypeType = Type.GetType(allowedContentTypeTypeName, true);
                var allowedContentTypesArray = Array.CreateInstance(allowedContentTypeType, documentSetTemplate.AllowedContentTypes.Count);

                result.GetPublicInstanceProperty("RemoveExistingContentTypes").SetValue(result, documentSetTemplate.RemoveExistingContentTypes);

                Int32 i = 0;
                foreach (var ct in documentSetTemplate.AllowedContentTypes)
                {
                    var item = Activator.CreateInstance(allowedContentTypeType);
                    item.SetPublicInstancePropertyValue("ContentTypeID", ct);
                    allowedContentTypesArray.SetValue(item, i);
                    i++;
                }

                if (allowedContentTypesArray.Length > 0)
                {
                    result.SetPublicInstancePropertyValue("AllowedContentType", allowedContentTypesArray);
                }
                else
                {
                    result = null;
                }
            }

            return (result);
        }
    }
}
