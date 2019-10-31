using System;
using System.ComponentModel;
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for tenant extension methods
    /// </summary>
    public static partial class TenantExtensions
    {
#if !ONPREMISES

        [Obsolete("Use ApplyTenantTemplate(this Tenant tenant, ProvisioningHierarchy tenantTemplate, string sequenceId, ApplyConfiguration configuration). This method will be removed in the May 2020 release.")]
        public static void ApplyProvisionHierarchy(this Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation = null)
        {
            if (applyingInformation == null)
            {
                ApplyTenantTemplate(tenant, hierarchy, sequenceId);
            }
            else
            {
                ApplyTenantTemplate(tenant, hierarchy, sequenceId, ApplyConfiguration.FromApplyingInformation(applyingInformation));
            }
        }
#endif
    }
}
