using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201909
{
    /// <summary>
    /// Type resolver for Office365GroupsSettings from Model to Schema
    /// </summary>
    internal class Office365GroupsSettingsFromModelToSchema : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Object result = null;

            // Declare supporting types
            var propertiesTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var propertiesType = Type.GetType(propertiesTypeName, true);

            var settings = ((Model.ProvisioningTenant)source).Office365GroupsSettings;

            if (null != settings &&
                settings.Properties != null && 
                settings.Properties.Count > 0)
            {
                var resultArray = Array.CreateInstance(propertiesType, settings.Properties.Count);

                int index = 0;
                foreach (var i in settings.Properties)
                {
                    var targetItem = Activator.CreateInstance(propertiesType, true);
                    targetItem.SetPublicInstancePropertyValue("Key", i.Key);
                    targetItem.SetPublicInstancePropertyValue("Value", i.Value);
                    resultArray.SetValue(targetItem, index++);
                }
                result = resultArray;
            }

            return (result);
        }
    }
}
