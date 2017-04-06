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
    /// Class to serialize/deserialize the Audit Settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        SerializationSequence = 600, DeserializationSequence = 600,
        Default = false)]
    internal class AuditSettingsSerializer : PnPBaseSchemaSerializer<AuditSettings>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var auditSettings = persistence.GetPublicInstancePropertyValue("AuditSettings");

            if (auditSettings != null)
            {
                var expressions = new Dictionary<Expression<Func<AuditSettings, Object>>, IResolver>();
                expressions.Add(a => a.AuditFlags, new FromArrayToAuditFlagsResolver());

                template.AuditSettings = new AuditSettings();
                PnPObjectsMapper.MapProperties(auditSettings, template.AuditSettings, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.AuditSettings != null)
            {
                var auditSettingsType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.AuditSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var target = Activator.CreateInstance(auditSettingsType, true);
                var expressions = new Dictionary<string, IResolver>();
                expressions.Add($"{auditSettingsType}.AuditLogTrimmingRetentionSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{auditSettingsType}.TrimAuditLogSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{auditSettingsType}.Audit", new FromAuditFlagsToArrayResolver());

                PnPObjectsMapper.MapProperties(template.AuditSettings, target, expressions, recursive: true);

                persistence.GetPublicInstanceProperty("AuditSettings").SetValue(persistence, target);
            }
        }
    }
}
