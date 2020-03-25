using System;
using System.Collections;
using System.Collections.Generic;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    /// <summary>
    /// Resolves a list of Shared Fields from Schema to Domain Model
    /// </summary>
    internal class DocumentSetTemplateSharedFieldsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public DocumentSetTemplateSharedFieldsFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new List<Model.SharedField>();

            var sharedFieldsContainer = source.GetPublicInstancePropertyValue("SharedFields");
            var sharedFieldTypes = sharedFieldsContainer?.GetPublicInstancePropertyValue("SharedField");

            if (null != sharedFieldTypes)
            {
                foreach(var field in (IEnumerable)sharedFieldTypes)
                {
                    var model = new Model.SharedField
                    {                        
                        Name = field?.GetPublicInstancePropertyValue("Name")?.ToString(),
                    };

                    var fieldId = field?.GetPublicInstancePropertyValue("ID");
                    if (fieldId != null && Guid.TryParse(fieldId.ToString(), out Guid fieldIdGuid))
                    {
                        model.FieldId = fieldIdGuid;
                    }

                    var removeSharedField = field?.GetPublicInstancePropertyValue("Remove");
                    if (removeSharedField != null && bool.TryParse(removeSharedField.ToString(), out bool removeSharedFieldBool))
                    {
                        model.Remove = removeSharedFieldBool;
                    }

                    result.Add(model);
                }
            }

            return (result);
        }
    }
}
