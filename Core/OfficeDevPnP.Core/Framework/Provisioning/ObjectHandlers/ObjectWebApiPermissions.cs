#if !ONPREMISES
using System;
using System.Linq;
using Microsoft.Online.SharePoint.TenantAdministration.Internal;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWebApiPermissions : ObjectHandlerBase
    {
        public override string Name => "Web API Permissions";

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (template.Tenant != null && template.Tenant.WebApiPermissions != null)
            {
                if (template.Tenant.WebApiPermissions.Any())
                {
                    using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl()))
                    {
                        var servicePrincipal = new SPOWebAppServicePrincipal(tenantContext);
                        //var requests = servicePrincipal.PermissionRequests;
                        var requestsEnumerable = tenantContext.LoadQuery(servicePrincipal.PermissionRequests);
                        var grantsEnumerable = tenantContext.LoadQuery(servicePrincipal.PermissionGrants);
                        tenantContext.ExecuteQueryRetry();

                        var requests = requestsEnumerable.ToList();

                        foreach (var permission in template.Tenant.WebApiPermissions)
                        {
                            var request = requests.FirstOrDefault(r => r.Scope.Equals(permission.Scope, StringComparison.InvariantCultureIgnoreCase) && r.Resource.Equals(permission.Resource, StringComparison.InvariantCultureIgnoreCase));
                            while (request != null)
                            {
                                if (grantsEnumerable.FirstOrDefault(g => g.Resource.Equals(permission.Resource, StringComparison.InvariantCultureIgnoreCase) && g.Scope.ToLower().Contains(permission.Scope.ToLower())) == null)
                                {
                                    var requestToApprove = servicePrincipal.PermissionRequests.GetById(request.Id);
                                    tenantContext.Load(requestToApprove);
                                    tenantContext.ExecuteQueryRetry();
                                    try
                                    {
                                        requestToApprove.Approve();
                                        tenantContext.ExecuteQueryRetry();
                                    }
                                    catch (Exception ex)
                                    {
                                        WriteMessage(ex.Message, ProvisioningMessageType.Warning);
                                    }
                                }
                                requests.Remove(request);
                                request = requests.FirstOrDefault(r => r.Scope.Equals(permission.Scope, StringComparison.InvariantCultureIgnoreCase) && r.Resource.Equals(permission.Resource, StringComparison.InvariantCultureIgnoreCase));
                            }
                        }
                    }
                }
            }
            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return false;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return (template.Tenant != null && template.Tenant.WebApiPermissions != null && template.Tenant.WebApiPermissions.Any());
        }
    }
}
#endif
