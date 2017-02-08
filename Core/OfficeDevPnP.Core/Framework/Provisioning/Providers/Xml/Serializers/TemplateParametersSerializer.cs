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
    [TemplateSchemaSerializer(
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate), typeof(Xml.V201512.ProvisioningTemplate) },
        Default = false)]
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

                if (parameters != null)
                {
                    template.GetType().GetProperty("Parameters")
                        .SetValue(template, PnPObjectsMapper.MapObject(parameters,
                                new TemplateParameterFromSchemaToModelTypeResolver())
                                as Dictionary<String, String>);
                }
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var preferences = persistence.GetType().GetProperty("Preferences",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.Public).GetValue(persistence);

            if (preferences != null)
            {
                var parametersTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PreferencesParameter, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var parametersType = Type.GetType(parametersTypeName, true);

                preferences.GetType().GetProperty("Parameters",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public).SetValue(
                        preferences,
                        PnPObjectsMapper.MapObject(template.Parameters,
                            new TemplateParameterFromModelToSchemaTypeResolver(parametersType)));
            }
        }
    }
}
