using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.ALM;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules.Rules
{
    [CanProvisionRule(Scope = CanProvisionScope.Site, Sequence = 100)]
    internal class CanProvisionAppsRuleSite : CanProvisionRuleSiteBase
    {
        public override CanProvisionResult CanProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // Prepare the default output
            var result = new CanProvisionResult();

#if !SP2013 && !SP2016

            Model.ProvisioningTemplate targetTemplate = null;

            if (template.ParentHierarchy != null)
            {
                // If we have a hierarchy, search for a template with ALM settings, if any
                targetTemplate = template.ParentHierarchy.Templates.FirstOrDefault(t => t.ApplicationLifecycleManagement.Apps.Count > 0 ||
                    (t.ApplicationLifecycleManagement.AppCatalog != null && t.ApplicationLifecycleManagement.AppCatalog.Packages.Count > 0));

                if (targetTemplate == null)
                {
                    // or use the first in the hierarchy
                    targetTemplate = template.ParentHierarchy.Templates[0];
                }
            }
            else
            {
                // Otherwise, use the provided template
                targetTemplate = template;
            }

            // Verify if we need the App Catalog (i.e. the template contains apps or packages)
            if ((targetTemplate.ApplicationLifecycleManagement?.Apps != null && targetTemplate.ApplicationLifecycleManagement?.Apps?.Count > 0) ||
                (targetTemplate.ApplicationLifecycleManagement?.AppCatalog != null &&
                targetTemplate.ApplicationLifecycleManagement?.AppCatalog?.Packages != null && targetTemplate.ApplicationLifecycleManagement?.AppCatalog?.Packages.Count > 0) ||
                (targetTemplate.ParentHierarchy != null && targetTemplate.ParentHierarchy?.Tenant?.AppCatalog != null &&
                targetTemplate.ParentHierarchy?.Tenant?.AppCatalog?.Packages != null && targetTemplate.ParentHierarchy?.Tenant?.AppCatalog?.Packages.Count > 0))
            {
                // First of all check if the currently connected user is a Tenant Admin
#if !ONPREMISES
                if (!TenantExtensions.IsCurrentUserTenantAdmin(web.Context as ClientContext))
#else
                if (!TenantExtensions.IsCurrentUserTenantAdmin(web.Context as ClientContext, this.TenantAdminSiteUrl))
#endif
                {
                    result.CanProvision = false;
                    result.Issues.Add(new CanProvisionIssue()
                    {
                        Source = this.Name,
                        Tag = CanProvisionIssueTags.USER_IS_NOT_TENANT_ADMIN,
                        Message = CanProvisionIssuesMessages.User_Is_Not_Tenant_Admin,
                        ExceptionMessage = null, // Here we don't have any specific exception
                        ExceptionStackTrace = null, // Here we don't have any specific exception
                    });
                }

                using (var scope = new PnPMonitoredScope(this.Name))
                {
                    // Try to access the AppCatalog
                    var appCatalogUri = web.GetAppCatalog();
                    if (appCatalogUri == null)
                    {
                        // And if we fail, raise a CanProvisionIssue
                        result.CanProvision = false;
                        result.Issues.Add(new CanProvisionIssue()
                        {
                            Source = this.Name,
                            Tag = CanProvisionIssueTags.MISSING_APP_CATALOG,
                            Message = CanProvisionIssuesMessages.Missing_App_Catalog,
                            ExceptionMessage = null, // Here we don't have any specific exception
                            ExceptionStackTrace = null, // Here we don't have any specific exception
                        });
                    }
                    else
                    {
                        // Try to access the AppCatalog with the current user

                        try
                        {
                            using (var appCatalogContext = web.Context.Clone(appCatalogUri))
                            {
                                // Get a reference to the "Apps for SharePoint" library
                                var appCatalogLibrary = appCatalogContext.Web.GetListByUrl("AppCatalog");

                                // Check its permissions
                                appCatalogContext.Web.CurrentUser.EnsureProperty(u => u.LoginName);
                                var userEffectivePermissions = appCatalogLibrary.GetUserEffectivePermissions(
                                    appCatalogContext.Web.CurrentUser.LoginName);
                                appCatalogContext.ExecuteQueryRetry();

                                if (!userEffectivePermissions.Value.Has(PermissionKind.EditListItems))
                                {
                                    throw new SecurityException("Invalid user's permissions for the AppCatalog");
                                }

                                // we seem to have access, but is it done fully provisioning?

                                var rootFolder = appCatalogContext.Web.EnsureProperty(w => w.RootFolder);
                                var timeCreated = rootFolder.TimeCreated;

                                if (DateTime.UtcNow.Subtract(timeCreated).TotalHours < 2)
                                {
                                    result.CanProvision = false;
                                    result.Issues.Add(new CanProvisionIssue()
                                    {
                                        Source = this.Name,
                                        Tag = CanProvisionIssueTags.APP_CATALOG_NOT_YEY_FULLY_PROVISIONED,
                                        Message = CanProvisionIssuesMessages.App_Catalog_Not_Yet_Fully_Provisioned,
                                        ExceptionMessage = null, // Here we don't have any specific exception
                                        ExceptionStackTrace = null, // Here we don't have any specific exception
                                    });
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            // And if we fail, raise a CanProvisionIssue
                            result.CanProvision = false;
                            result.Issues.Add(new CanProvisionIssue()
                            {
                                Source = this.Name,
                                Tag = CanProvisionIssueTags.MISSING_APP_CATALOG_PERMISSIONS,
                                Message = CanProvisionIssuesMessages.Missing_Permissions_for_App_Catalog,
                                ExceptionMessage = ex.Message,
                                ExceptionStackTrace = ex.StackTrace,
                            });
                        }
                    }
                }
            }
#else
            result.CanProvision = false;
#endif
                    return result;
        }
    }
}
