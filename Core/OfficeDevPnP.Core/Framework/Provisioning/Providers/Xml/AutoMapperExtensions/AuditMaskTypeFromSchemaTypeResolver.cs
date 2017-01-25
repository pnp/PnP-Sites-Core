using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class AuditMaskTypeFromSchemaTypeResolver : IMemberValueResolver<object, object, V201605.AuditSettingsAudit[], Microsoft.SharePoint.Client.AuditMaskType>
    {
        public Microsoft.SharePoint.Client.AuditMaskType Resolve(object source, object destination, V201605.AuditSettingsAudit[] sourceMember, Microsoft.SharePoint.Client.AuditMaskType destMember, ResolutionContext context)
        {
            return sourceMember.Aggregate(Microsoft.SharePoint.Client.AuditMaskType.None, (acc, next) => acc |= (Microsoft.SharePoint.Client.AuditMaskType) Enum.Parse(typeof(Microsoft.SharePoint.Client.AuditMaskType), next.AuditFlag.ToString()));
        }
    }
}