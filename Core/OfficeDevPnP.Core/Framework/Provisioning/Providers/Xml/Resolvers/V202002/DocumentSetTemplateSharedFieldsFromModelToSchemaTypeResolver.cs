using System;
using System.Collections.Generic;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class DocumentSetTemplateSharedFieldsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public DocumentSetTemplateSharedFieldsFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            var documentSetTemplate = source as Model.DocumentSetTemplate;

            if (null != documentSetTemplate)
            {
                var sharedFieldsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DocumentSetTemplateSharedFields, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var sharedFieldsType = Type.GetType(sharedFieldsTypeName, true);
                result = Activator.CreateInstance(sharedFieldsType);

                var sharedFieldTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DocumentSetTemplateSharedFieldsSharedField, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var sharedFieldType = Type.GetType(sharedFieldTypeName, true);
                var sharedFieldsArray = Array.CreateInstance(sharedFieldType, documentSetTemplate.SharedFields.Count);

                Int32 i = 0;
                foreach (var field in documentSetTemplate.SharedFields)
                {
                    var item = Activator.CreateInstance(sharedFieldType);
                    item.SetPublicInstancePropertyValue("ID", field.FieldId);
                    item.SetPublicInstancePropertyValue("Name", field.Name);
                    item.SetPublicInstancePropertyValue("Remove", field.Remove);
                    sharedFieldsArray.SetValue(item, i);
                    i++;
                }

                if (sharedFieldsArray.Length > 0)
                {
                    result.SetPublicInstancePropertyValue("SharedField", sharedFieldsArray);
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
