using System;
using System.Collections;
using System.Collections.Generic;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    /// <summary>
    /// Resolves a list of Shared Fields from Schema to Domain Model
    /// </summary>
    internal class DocumentSetTemplateWelcomePageFieldsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public DocumentSetTemplateWelcomePageFieldsFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new List<Model.WelcomePageField>();

            var welcomePageFieldsContainer = source.GetPublicInstancePropertyValue("WelcomePageFields");
            var welcomePageFieldTypes = welcomePageFieldsContainer?.GetPublicInstancePropertyValue("WelcomePageField");

            if (null != welcomePageFieldTypes)
            {
                foreach(var field in (IEnumerable)welcomePageFieldTypes)
                {
                    var model = new Model.WelcomePageField
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
