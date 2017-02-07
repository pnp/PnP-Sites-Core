using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the parameters of the template
    /// </summary>
    internal class TemplateParametersSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var preferences = persistence.GetType().GetProperty("Preferences",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.Public).GetValue(persistence);
            
            if (preferences != null)
            {
                var parameters = preferences.GetType().GetProperty("Parameters",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public).GetValue(preferences);

                template.GetType().GetProperty("Parameters")
                    .SetValue(template, PnPObjectsMapper.MapObject(parameters,
                            new TemplateParameterFromSchemaToModelTypeResolver())
                            as Dictionary<String, String>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var parametersTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PreferencesParameter, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var parametersType = Type.GetType(parametersTypeName, true);

            persistence.GetType().GetProperty("Parameters",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.Public).SetValue(
                    persistence,
                    PnPObjectsMapper.MapObject(template.Parameters,
                        new TemplateParameterFromModelToSchemaTypeResolver(parametersType)));
        }
    }
}
