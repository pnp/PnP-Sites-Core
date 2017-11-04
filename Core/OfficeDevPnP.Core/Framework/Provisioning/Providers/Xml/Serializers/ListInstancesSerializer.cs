using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the List Instances
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 1100, DeserializationSequence = 1100,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class ListInstancesSerializer : PnPBaseSchemaSerializer<ListInstance>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var lists = persistence.GetPublicInstancePropertyValue("Lists");

            if (lists != null)
            {
                var expressions = new Dictionary<Expression<Func<ListInstance, Object>>, IResolver>();

                // Define custom resolver for FieldRef.ID because needs conversion from String to GUID
                expressions.Add(l => l.FieldRefs[0].Id, new FromStringToGuidValueResolver());
                expressions.Add(l => l.TemplateFeatureID, new FromStringToGuidValueResolver());

                // Define custom resolvers for DataRows Values and Security
                var dataRowValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DataValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dataRowValueType = Type.GetType(dataRowValueTypeName, true);
                var dataRowValueKeySelector = CreateSelectorLambda(dataRowValueType, "FieldName");
                var dataRowValueValueSelector = CreateSelectorLambda(dataRowValueType, "Value");
                expressions.Add(l => l.DataRows[0].Values,
                    new FromArrayToDictionaryValueResolver<String, String>(
                        dataRowValueType, dataRowValueKeySelector, dataRowValueValueSelector));

                expressions.Add(l => l.DataRows[0].Security, new SecurityFromSchemaToModelTypeResolver());

                // Define custom resolver for Fields Defaults
                var fieldDefaultTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.FieldDefault, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var fieldDefaultType = Type.GetType(fieldDefaultTypeName, true);
                var fieldDefaultKeySelector = CreateSelectorLambda(fieldDefaultType, "FieldName");
                var fieldDefaultValueSelector = CreateSelectorLambda(fieldDefaultType, "Value");
                expressions.Add(l => l.FieldDefaults,
                    new FromArrayToDictionaryValueResolver<String, String>(
                        fieldDefaultType, fieldDefaultKeySelector, fieldDefaultValueSelector));

                // Define custom resolver for Security
                expressions.Add(l => l.Security, new SecurityFromSchemaToModelTypeResolver());

                // Define custom resolver for UserCustomActions > CommandUIExtension (XML Any)
                expressions.Add(l => l.UserCustomActions[0].CommandUIExtension, new XmlAnyFromSchemaToModelValueResolver("CommandUIExtension"));
                expressions.Add(l => l.UserCustomActions[0].RegistrationType, new FromStringToEnumValueResolver(typeof(UserCustomActionRegistrationType)));
                expressions.Add(l => l.UserCustomActions[0].Rights, new FromStringToBasePermissionsValueResolver());

                // Define custom resolver for Views (XML Any + RemoveExistingViews)
                expressions.Add(l => l.Views,
                    new ListViewsFromSchemaToModelTypeResolver());
                expressions.Add(l => l.RemoveExistingViews,
                    new RemoveExistingViewsFromSchemaToModelValueResolver());

                // Define custom resolver for recursive Folders
                expressions.Add(l => l.Folders,
                   new FoldersFromSchemaToModelTypeResolver());

                // Fields
                expressions.Add(l => l.Fields, new ExpressionValueResolver((s, v) => {
                    var fields = new Model.FieldCollection(template);
                    var sourceFields = s.GetPublicInstancePropertyValue("Fields")?.GetPublicInstancePropertyValue("Any") as System.Xml.XmlElement[];
                    if (sourceFields != null)
                    {
                        foreach (var f in sourceFields)
                        {
                            fields.Add(new Model.Field { SchemaXml = f.OuterXml });
                        }
                    }
                    return fields;
                }));

                template.Lists.AddRange(
                    PnPObjectsMapper.MapObjects<ListInstance>(lists,
                            new CollectionFromSchemaToModelTypeResolver(typeof(ListInstance)),
                            expressions,
                            recursive: true)
                            as IEnumerable<ListInstance>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Lists != null && template.Lists.Count > 0)
            {
                var listInstanceTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstance, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var listInstanceType = Type.GetType(listInstanceTypeName, true);

                var resolvers = new Dictionary<String, IResolver>();

                // Define custom resolvers for DataRows Values and Security
                var listInstanceDataRowTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstanceDataRow, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var listInstanceDataRowType = Type.GetType(listInstanceDataRowTypeName, true);
                var listInstanceDataRowValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DataValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var listInstanceDataRowValueType = Type.GetType(listInstanceDataRowValueTypeName, true);

                var dataRowValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DataValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dataRowValueType = Type.GetType(dataRowValueTypeName, true);
                var dataRowValueKeySelector = CreateSelectorLambda(dataRowValueType, "FieldName");
                var dataRowValueValueSelector = CreateSelectorLambda(dataRowValueType, "Value");

                resolvers.Add($"{listInstanceDataRowType}.DataValue", new FromDictionaryToArrayValueResolver<string, string>(dataRowValueType, dataRowValueKeySelector, dataRowValueValueSelector, "Values"));
                resolvers.Add($"{listInstanceDataRowType}.Security", new SecurityFromModelToSchemaTypeResolver());

                // Define custom resolver for Fields Defaults
                var fieldDefaultTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.FieldDefault, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var fieldDefaultType = Type.GetType(fieldDefaultTypeName, true);
                var fieldDefaultKeySelector = CreateSelectorLambda(fieldDefaultType, "FieldName");
                var fieldDefaultValueSelector = CreateSelectorLambda(fieldDefaultType, "Value");

                resolvers.Add($"{listInstanceType}.FieldDefaults", new FromDictionaryToArrayValueResolver<string, string>(fieldDefaultType, fieldDefaultKeySelector, fieldDefaultValueSelector));

                // Define custom resolver for Security
                resolvers.Add($"{listInstanceType}.Security", new SecurityFromModelToSchemaTypeResolver());

                // Define custom resolver for UserCustomActions > CommandUIExtension (XML Any)
                var customActionTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CustomAction, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var customActionType = Type.GetType(customActionTypeName, true);
                var commandUIExtensionTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CustomActionCommandUIExtension";
                var commandUIExtensionType = Type.GetType(commandUIExtensionTypeName, true);
                var registrationTypeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.RegistrationType";
                var registrationTypeType = Type.GetType(registrationTypeTypeName, true);
                resolvers.Add($"{customActionType}.CommandUIExtension", new XmlAnyFromModelToSchemalValueResolver(commandUIExtensionType));
                resolvers.Add($"{customActionType}.Rights", new FromBasePermissionsToStringValueResolver());
                resolvers.Add($"{customActionType}.RegistrationType", new FromStringToEnumValueResolver(registrationTypeType));
                resolvers.Add($"{customActionType}.RegistrationTypeSpecified", new ExpressionValueResolver(() => true));
                resolvers.Add($"{customActionType}.SequenceSpecified", new ExpressionValueResolver(() => true));


                // Define custom resolver for Views (XML Any + RemoveExistingViews)
                var listInstanceViewsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstanceViews, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var listInstanceViewsType = Type.GetType(listInstanceViewsTypeName, true);

                resolvers.Add($"{listInstanceType}.Views",
                    new ListViewsFromModelToSchemaTypeResolver());
                resolvers.Add($"{listInstanceViewsType}.RemoveExistingViews",
                    new ExpressionValueResolver((s, v) => (Boolean)s.GetPublicInstancePropertyValue("RemoveExistingViews")));

                // Define custom resolver for recursive Folders
                resolvers.Add($"{listInstanceType}.Folders", new FoldersFromModelToSchemaTypeResolver());

                // Fields
                var fieldsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstanceFields, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var fieldsType = Type.GetType(fieldsTypeName, true);
                resolvers.Add($"{listInstanceType}.Fields", new ExpressionValueResolver<ListInstance>((s, v) =>
                {
                    if (s.Fields != null && s.Fields.Count > 0)
                    {
                        var fields = Activator.CreateInstance(fieldsType);
                        var xmlFields = from f in s.Fields
                                        select XElement.Parse(f.SchemaXml).ToXmlElement();

                        fields.SetPublicInstancePropertyValue("Any", xmlFields.ToArray());
                        return fields;
                    }
                    else
                    {
                        return null;
                    }
                }));

                resolvers.Add($"{listInstanceType}.DraftVersionVisibilitySpecified", new ExpressionValueResolver(() => true));
                resolvers.Add($"{listInstanceType}.MaxVersionLimitSpecified", new ExpressionValueResolver(() => true));
                resolvers.Add($"{listInstanceType}.MinorVersionLimitSpecified", new ExpressionValueResolver(() => true));

                persistence.GetPublicInstanceProperty("Lists")
                    .SetValue(
                        persistence,
                        PnPObjectsMapper.MapObjects(template.Lists,
                            new CollectionFromModelToSchemaTypeResolver(listInstanceType), resolvers, recursive: true));
            }
        }
    }
}
