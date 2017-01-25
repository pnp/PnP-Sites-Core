using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class AuditMaskTypeFromModelTypeResolver : IMemberValueResolver<object, object, Microsoft.SharePoint.Client.AuditMaskType, V201605.AuditSettingsAudit[]>
    {
        public AuditSettingsAudit[] Resolve(object source, object destination, AuditMaskType sourceMember, AuditSettingsAudit[] destMember, ResolutionContext context)
        {
            var audits = sourceMember;
            List<V201605.AuditSettingsAudit> result = new List<V201605.AuditSettingsAudit>();
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.All))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.All });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.CheckIn))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.CheckIn });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.CheckOut))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.CheckOut });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.ChildDelete))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.ChildDelete });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Copy))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Copy });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Move))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Move });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.None))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.None });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.ObjectDelete))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.ObjectDelete });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.ProfileChange))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.ProfileChange });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.SchemaChange))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.SchemaChange });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Search))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Search });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.SecurityChange))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.SecurityChange });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Undelete))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Undelete });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Update))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Update });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.View))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.View });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Workflow))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Workflow });
            }

            return result.ToArray();
        }
    }
}
