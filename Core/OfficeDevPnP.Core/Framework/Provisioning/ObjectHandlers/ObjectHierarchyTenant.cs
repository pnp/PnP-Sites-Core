#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectHierarchyTenant : ObjectHierarchyHandlerBase
    {
        public override string Name
        {
            get { return "Tenant Settings"; }
        }

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ProvisioningTemplateCreationInformation creationInfo)
        {
            return hierarchy;
        }

        public override TokenParser ProvisionObjects(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (hierarchy.Tenant != null)
                {
                    TenantHelper.ProcessCdns(tenant, hierarchy.Tenant, parser, scope, MessagesDelegate);
                    parser = TenantHelper.ProcessApps(tenant, hierarchy.Tenant, hierarchy.Connector, parser, scope, applyingInformation, MessagesDelegate);

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
                    parser = TenantHelper.ProcessStorageEntities(tenant, hierarchy.Tenant, parser, scope, applyingInformation, MessagesDelegate);
                    parser = TenantHelper.ProcessThemes(tenant, hierarchy.Tenant, parser, scope, MessagesDelegate);
                    // So far we do not provision CDN settings
                    // It will come in the near future
                    // NOOP on CDN
                }
            }

            return parser;
        }

        public override bool WillExtract(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateCreationInformation creationInfo)
        {
            return false;
        }

        public override bool WillProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return hierarchy.Tenant != null;
        }
    }
}
#endif
