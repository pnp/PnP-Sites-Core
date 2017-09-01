using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Publishing settings
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 1900, DeserializationSequence = 1900,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class PublishingSerializer : PnPBaseSchemaSerializer<Publishing>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var publishing = persistence.GetPublicInstancePropertyValue("Publishing");

            if (publishing != null)
            {
                template.Publishing = new Publishing();
                var expressions = new Dictionary<Expression<Func<Publishing, Object>>, IResolver>();

                expressions.Add(p => p.DesignPackage, new PropertyObjectTypeResolver<Publishing>(p => p.DesignPackage));
                expressions.Add(p => p.DesignPackage.PackageGuid, new FromStringToGuidValueResolver());
                expressions.Add(p => p.PageLayouts, new PageLayoutsFromSchemaToModelTypeResolver());
                expressions.Add(p => p.AvailableWebTemplates[0].LanguageCode,
                    new ExpressionValueResolver((s, v) => (bool)s.GetPublicInstancePropertyValue("LanguageCodeSpecified") ? v : 1033));

                PnPObjectsMapper.MapProperties(publishing, template.Publishing, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Publishing != null)
            {
                var publishingType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Publishing, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var designPackageType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PublishingDesignPackage, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var webTemplateType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PublishingWebTemplate, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);

                var target = Activator.CreateInstance(publishingType, true);
                var expressions = new Dictionary<string, IResolver>();

                expressions.Add($"{publishingType}.DesignPackage", new PropertyObjectTypeResolver(designPackageType, "DesignPackage"));
                expressions.Add($"{designPackageType}.MajorVersionSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{designPackageType}.MinorVersionSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{webTemplateType}.LanguageCodeSpecified", new ExpressionValueResolver(() => true));
                expressions.Add($"{publishingType}.PageLayouts", new PageLayoutsFromModelToSchemaTypeResolver());

                PnPObjectsMapper.MapProperties(template.Publishing, target, expressions, recursive: true);

                persistence.GetPublicInstanceProperty("Publishing").SetValue(persistence, target);
            }
        }
    }
}
