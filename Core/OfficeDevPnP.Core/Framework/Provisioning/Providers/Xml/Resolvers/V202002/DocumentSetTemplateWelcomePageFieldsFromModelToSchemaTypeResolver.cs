using System;
using System.Collections.Generic;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class DocumentSetTemplateWelcomePageFieldsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public DocumentSetTemplateWelcomePageFieldsFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            var documentSetTemplate = source as Model.DocumentSetTemplate;

            if (null != documentSetTemplate)
            {
                var welcomePageFieldsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DocumentSetTemplateWelcomePageFields, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var welcomePageFieldsType = Type.GetType(welcomePageFieldsTypeName, true);
                result = Activator.CreateInstance(welcomePageFieldsType);

                var welcomePageFieldTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DocumentSetTemplateWelcomePageFieldsWelcomePageField, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var welcomePageFieldType = Type.GetType(welcomePageFieldTypeName, true);
                var welcomePageFieldsArray = Array.CreateInstance(welcomePageFieldType, documentSetTemplate.WelcomePageFields.Count);

                Int32 i = 0;
                foreach (var field in documentSetTemplate.SharedFields)
                {
                    var item = Activator.CreateInstance(welcomePageFieldType);
                    item.SetPublicInstancePropertyValue("ID", field.FieldId);
                    item.SetPublicInstancePropertyValue("Name", field.Name);
                    item.SetPublicInstancePropertyValue("Remove", field.Remove);
                    welcomePageFieldsArray.SetValue(item, i);
                    i++;
                }

                if (welcomePageFieldsArray.Length > 0)
                {
                    result.SetPublicInstancePropertyValue("WelcomePageFields", welcomePageFieldsArray);
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
