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

        /// <summary>
        /// Checks if a site collection exists, relies on tenant admin API. Sites that are recycled also return as existing sites
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">URL to the site collection</param>
        /// <returns>True if existing, false if not</returns>
        [Obsolete()]
        public static bool SiteExists(this Tenant tenant, string siteFullUrl)
        {
            var exists = SiteExistsAnywhere(tenant, siteFullUrl);
            return (exists == SiteExistence.Yes || exists == SiteExistence.Recycled);
        }
    }
}
