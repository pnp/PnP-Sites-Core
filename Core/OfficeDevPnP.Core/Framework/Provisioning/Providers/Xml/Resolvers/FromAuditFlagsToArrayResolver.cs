using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves an enum bit mask of AuditFlags into an array of Strings 
    /// </summary>
    internal class FromAuditFlagsToArrayResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        public object Resolve(object source, object destination, object sourceValue)
        {
            var auditValues = new List<AuditMaskType>();
            AuditMaskType auditFlags = (AuditMaskType)source.GetPublicInstancePropertyValue("AuditFlags");
            foreach (var f in Enum.GetValues(typeof(AuditMaskType)))
            {
                if (auditFlags.HasFlag((AuditMaskType)f) && ((AuditMaskType)f != AuditMaskType.None))
                {
                    auditValues.Add((AuditMaskType)f);
                }
            }

            var auditEnumType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.AuditSettingsAuditAuditFlag, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
            var auditType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.AuditSettingsAudit, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
            var target = Array.CreateInstance(auditType, auditValues.Count);

            for (Int32 c = 0; c < auditValues.Count; c++)
            {
                var targetAuditValue = Enum.Parse(auditEnumType, 
                    Enum.GetName(typeof(AuditMaskType), auditValues[c]));

                var targetAudit = Activator.CreateInstance(auditType, true);
                targetAudit.SetPublicInstancePropertyValue("AuditFlag", targetAuditValue);

                ((Array)target).SetValue(targetAudit, c);
            }

            return target;
        }
    }
}
