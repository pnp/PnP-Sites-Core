using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the AddIns
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 2000, DeserializationSequence = 2000,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class AddInsSerializer : PnPBaseSchemaSerializer<AddIn>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var addIns = persistence.GetPublicInstancePropertyValue("AddIns");

            if (addIns != null)
            {
                template.AddIns.AddRange(
                    PnPObjectsMapper.MapObjects<AddIn>(addIns,
                        new CollectionFromSchemaToModelTypeResolver(typeof(AddIn)),
                        null,
                        recursive: true)
                        as IEnumerable<AddIn>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.AddIns != null && template.AddIns.Count > 0)
            {
                var addInType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.AddInsAddin, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", false);
                var addinSourceType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.AddInsAddinSource, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", false);

                if (addInType != null && addinSourceType != null)
                {
                    var expressions = new Dictionary<string, IResolver>();
                    expressions.Add($"{addInType}.Source", new FromStringToEnumValueResolver(addinSourceType));

                    persistence.GetPublicInstanceProperty("Addins").SetValue(
                        persistence,
                        PnPObjectsMapper.MapObjects(template.AddIns, new CollectionFromModelToSchemaTypeResolver(addInType), expressions, true));
                }
            }
        }
    }
}
