using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the content types
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
            expressions.Add(c => c.DocumentTemplate, new ExpressionValueResolver((s,v) => v.GetPublicInstancePropertyValue("TargetName")));
            expressions.Add(c => c.DocumentSetTemplate, new PropertyObjectTypeResolver<ContentType>(ct => ct.DocumentSetTemplate));
            expressions.Add(c => c.DocumentSetTemplate.AllowedContentTypes, new ExpressionCollectionValueResolver<string>((s) => s.GetPublicInstancePropertyValue("ContentTypeID").ToString()));
            expressions.Add(c => c.DocumentSetTemplate.SharedFields, new ExpressionCollectionValueResolver<Guid>((s) => Guid.Parse(s.GetPublicInstancePropertyValue("ID").ToString())));

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
            var documentSetTemplateTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DocumentSetTemplate, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var documentSetTemplateType = Type.GetType(documentSetTemplateTypeName, true);
            var documentTemplateTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ContentTypeDocumentTemplate, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var documentTemplateType = Type.GetType(documentTemplateTypeName, true);
            var expressions = new Dictionary<Expression<Func<Xml.V201605.ContentType, Object>>, IResolver>();

            expressions.Add(c => c.DocumentSetTemplate, new PropertyObjectTypeResolver(documentSetTemplateType, "DocumentSetTemplate"));
            expressions.Add(c => c.DocumentSetTemplate.AllowedContentTypes[0].ContentTypeID, new ExpressionValueResolver((s, v) => s));
            //this expression also used to resolve fieldref collection ids
            expressions.Add(c => c.DocumentSetTemplate.SharedFields[0].ID, new ExpressionValueResolver((s, v) => v != null ? v.ToString() : s?.ToString()));

            expressions.Add(c => c.DocumentTemplate, new ExpressionTypeResolver<ContentType>(documentTemplateType, 
                (s, r) => { r.SetPublicInstancePropertyValue("TargetName", s.DocumentTemplate); }));

            persistence.GetPublicInstanceProperty("ContentTypes")
                .SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.ContentTypes,
                        new CollectionFromModelToSchemaTypeResolver(contentTypeType), expressions, true));
        }
    }
}
