#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectHierarchyTenant : ObjectHierarchyHandlerBase
    {
        public override string Name
        {
            get { return "Tenant Settings"; }
        }

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ExtractConfiguration configuration)
        {
            return hierarchy;
        }

        public override TokenParser ProvisionObjects(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser, ApplyConfiguration configuration)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (hierarchy.Tenant != null)
                {
                    TenantHelper.ProcessCdns(tenant, hierarchy.Tenant, parser, scope, MessagesDelegate);
                    parser = TenantHelper.ProcessApps(tenant, hierarchy.Tenant, hierarchy.Connector, parser, scope, configuration, MessagesDelegate);

                    try
                    {
                        parser = TenantHelper.ProcessWebApiPermissions(tenant, hierarchy.Tenant, parser, scope, MessagesDelegate);
                    }
                    catch (ServerUnauthorizedAccessException ex)
                    {
                        scope.LogError(ex.Message);
                    }

                    parser = TenantHelper.ProcessSiteScripts(tenant, hierarchy.Tenant, hierarchy.Connector, parser, scope, MessagesDelegate);
                    parser = TenantHelper.ProcessSiteDesigns(tenant, hierarchy.Tenant, parser, scope, MessagesDelegate);
                    parser = TenantHelper.ProcessStorageEntities(tenant, hierarchy.Tenant, parser, scope, configuration, MessagesDelegate);
                    parser = TenantHelper.ProcessThemes(tenant, hierarchy.Tenant, parser, scope, MessagesDelegate);
                    parser = TenantHelper.ProcessUserProfiles(tenant, hierarchy.Tenant, parser, scope, MessagesDelegate);
                    // So far we do not provision CDN settings
                    // It will come in the near future
                    // NOOP on CDN
                }
            }

            return parser;
        }

        public override bool WillExtract(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ExtractConfiguration configuration)
        {
            return false;
        }

        public override bool WillProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration)
        {
            if (!_willProvision.HasValue && hierarchy.Tenant != null)
            {
                _willProvision = ((hierarchy.Tenant.AppCatalog?.Packages != null && hierarchy.Tenant.AppCatalog?.Packages.Count > 0 ) ||
                                hierarchy.Tenant.ContentDeliveryNetwork != null ||
                                (hierarchy.Tenant.SiteDesigns != null && hierarchy.Tenant.SiteDesigns.Count > 0) ||
                                (hierarchy.Tenant.SiteScripts != null && hierarchy.Tenant.SiteScripts.Count > 0) ||
                                (hierarchy.Tenant.StorageEntities != null && hierarchy.Tenant.StorageEntities.Count > 0) ||
                                (hierarchy.Tenant.WebApiPermissions != null && hierarchy.Tenant.WebApiPermissions.Count > 0) ||
                                (hierarchy.Tenant.Themes != null && hierarchy.Tenant.Themes.Count > 0) ||
                                (hierarchy.Tenant.SPUsersProfiles != null && hierarchy.Tenant.SPUsersProfiles.Count > 0) ||
                                (hierarchy.Tenant.Office365GroupLifecyclePolicies != null && hierarchy.Tenant.Office365GroupLifecyclePolicies.Count > 0) ||
                                (hierarchy.Tenant.Office365GroupsSettings?.Properties != null && hierarchy.Tenant.Office365GroupsSettings?.Properties.Count > 0)
                                );
            }
            return (_willProvision.Value);
        }
    }
}
#endif
