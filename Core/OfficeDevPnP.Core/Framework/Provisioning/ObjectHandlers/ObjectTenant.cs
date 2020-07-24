using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
#if !ONPREMISES
    internal class ObjectTenant : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Tenant Settings"; }
        }

        public override string InternalName => "TenantSettings";

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            web.EnsureProperty(w => w.Url);

            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Tenant != null)
                {
                    using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl()))
                    {
                        var tenant = new Tenant(tenantContext);
                        TenantHelper.ProcessCdns(tenant, template.Tenant, parser, scope, MessagesDelegate);
                        parser = TenantHelper.ProcessApps(tenant, template.Tenant, template.Connector, parser, scope, ApplyConfiguration.FromApplyingInformation(applyingInformation), MessagesDelegate);

                        try
                        {
                            parser = TenantHelper.ProcessWebApiPermissions(tenant, template.Tenant, parser, scope, MessagesDelegate);
                        }
                        catch (ServerUnauthorizedAccessException ex)
                        {
                            scope.LogError(ex.Message);
                        }

                        parser = TenantHelper.ProcessSiteScripts(tenant, template.Tenant, template.Connector, parser, scope, MessagesDelegate);
                        parser = TenantHelper.ProcessSiteDesigns(tenant, template.Tenant, parser, scope, MessagesDelegate);
                        parser = TenantHelper.ProcessStorageEntities(tenant, template.Tenant, parser, scope, ApplyConfiguration.FromApplyingInformation(applyingInformation), MessagesDelegate);
                        parser = TenantHelper.ProcessThemes(tenant, template.Tenant, parser, scope, MessagesDelegate);
                        parser = TenantHelper.ProcessUserProfiles(tenant, template.Tenant, parser, scope, MessagesDelegate);
                        parser = TenantHelper.ProcessSharingSettings(tenant, template.Tenant, parser, scope, MessagesDelegate);
                    }
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            // By default we don't extract the tenant settings
            return false;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue && template.Tenant != null)
            {
                _willProvision = (template.Tenant.AppCatalog != null ||
                                template.Tenant.ContentDeliveryNetwork != null ||
                                (template.Tenant.SiteDesigns != null && template.Tenant.SiteDesigns.Count > 0) ||
                                (template.Tenant.SiteScripts!= null && template.Tenant.SiteScripts.Count > 0) ||
                                (template.Tenant.StorageEntities != null && template.Tenant.StorageEntities.Count > 0) ||
                                (template.Tenant.WebApiPermissions!= null && template.Tenant.WebApiPermissions.Count > 0) ||
                                (template.Tenant.Themes != null && template.Tenant.Themes.Count > 0) ||
                                template.Tenant.SharingSettings != null
                                );
            }
            return (_willProvision.Value);
        }
    }

#endif
}
