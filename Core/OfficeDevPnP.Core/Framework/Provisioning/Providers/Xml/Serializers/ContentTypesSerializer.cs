using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the list instances
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 250, DeserializationSequence = 250,
        SchemaTemplates = new Type[] { typeof(Xml.V201605.ProvisioningTemplate) },
        Default = true)]
    internal class ContentTypesSerializer : PnPBaseSchemaSerializer
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var contentTypes = persistence.GetPublicInstancePropertyValue("ContentTypes");

            var expressions = new Dictionary<Expression<Func<ContentType, Object>>, IResolver>();

            // Define custom resolver for FieldRef.ID because needs conversion from String to GUID
            expressions.Add(c => c.FieldRefs[0].Id, new FromStringToGuidValueResolver());

            // Define custom resolver for Security
            //expressions.Add(l => l.DocumentSetTemplate, new SecurityFromSchemaToModelTypeResolver());

            template.ContentTypes.AddRange(
                PnPObjectsMapper.MapObjects<ContentType>(contentTypes,
                        new CollectionFromSchemaToModelTypeResolver(typeof(ContentType)),
                        expressions,
                        recursive: true)
                        as IEnumerable<ContentType>);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var contentTypeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ContentType, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var contentTypeType = Type.GetType(contentTypeTypeName, true);

            // TODO: Define the way back for the non-standard properties defined in the Deserialize method

            persistence.GetPublicInstanceProperty("ContentType")
                .SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.ContentTypes,
                        new CollectionFromModelToSchemaTypeResolver(contentTypeType), recursive: true));
        }
    }
}
