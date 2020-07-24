using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using OfficeDevPnP.Core.Utilities;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSiteSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Site Settings"; }
        }

        public override string InternalName => "SiteSettings";
        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
#if !ONPREMISES
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Try to get access to the Site Collection in the current context
                var site = (web.Context as ClientContext)?.Site;
                if (site != null)
                {
                    // And if we have it, load the properties in which we're interested into
                    site.EnsureProperties(
                        s => s.AllowDesigner,
                        s => s.AllowCreateDeclarativeWorkflow,
                        s => s.AllowSaveDeclarativeWorkflowAsTemplate,
                        s => s.AllowSavePublishDeclarativeWorkflow,
                        s => s.SocialBarOnSitePagesDisabled,
                        s => s.SearchBoxInNavBar
                        );

                    // Configure the output SiteSettings object
                    var siteSettings = new SiteSettings();

                    siteSettings.AllowDesigner = site.AllowDesigner;
                    siteSettings.AllowCreateDeclarativeWorkflow = site.AllowCreateDeclarativeWorkflow;
                    siteSettings.AllowSaveDeclarativeWorkflowAsTemplate = site.AllowSaveDeclarativeWorkflowAsTemplate;
                    siteSettings.AllowSavePublishDeclarativeWorkflow = site.AllowSavePublishDeclarativeWorkflow;
                    siteSettings.SocialBarOnSitePagesDisabled = site.SocialBarOnSitePagesDisabled;
                    siteSettings.SearchBoxInNavBar = (SearchBoxInNavBar)Enum.Parse(typeof(SearchBoxInNavBar), site.SearchBoxInNavBar.ToString());
                    siteSettings.SearchCenterUrl = site.RootWeb.GetSiteCollectionSearchCenterUrl();

                    // Update the provisioning template accordingly
                    template.SiteSettings = siteSettings;
                }
            }
#endif
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#if !ONPREMISES
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.SiteSettings != null)
                {
                    // Try to get access to the Site Collection in the current context
                    var site = (web.Context as ClientContext)?.Site;
                    if (site != null)
                    {
                        bool isDirty = false;

                        // Apply the following properties if and only if the target site is a classic one
                        if (!(site.IsCommunicationSite() || site.IsModernTeamSite()))
                        {
                            site.AllowDesigner = template.SiteSettings.AllowDesigner;
                            site.AllowCreateDeclarativeWorkflow = template.SiteSettings.AllowCreateDeclarativeWorkflow;
                            site.AllowSaveDeclarativeWorkflowAsTemplate = template.SiteSettings.AllowSaveDeclarativeWorkflowAsTemplate;
                            site.AllowSavePublishDeclarativeWorkflow = template.SiteSettings.AllowSavePublishDeclarativeWorkflow;
                            isDirty = true;
                        }

                        // Right now this works in Communication Sites only
                        // For further details: https://github.com/SharePoint/sp-dev-docs/issues/1532
                        if (site.IsCommunicationSite())
                        {
                            site.SocialBarOnSitePagesDisabled = template.SiteSettings.SocialBarOnSitePagesDisabled;
                            isDirty = true;
                        }

                        site.EnsureProperty(s => s.SearchBoxInNavBar);
                        if (site.SearchBoxInNavBar.ToString() != template.SiteSettings.SearchBoxInNavBar.ToString())
                        {
                            site.SearchBoxInNavBar = (SearchBoxInNavBarType)Enum.Parse(typeof(SearchBoxInNavBarType), template.SiteSettings.SearchBoxInNavBar.ToString(), true);
                            isDirty = true;
                        }

                        if (!string.IsNullOrEmpty(template.SiteSettings.SearchCenterUrl) &&
                            site.RootWeb.GetSiteCollectionSearchCenterUrl() != template.SiteSettings.SearchCenterUrl)
                        {
                            site.RootWeb.SetSiteCollectionSearchCenterUrl(template.SiteSettings.SearchCenterUrl);
                        }

                        // And save on SharePoint, if really needed
                        if (isDirty)
                        {
                            web.Context.ExecuteQueryRetryAsync();
                        }
                    }
                }
            }
#endif
            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return template.SiteSettings != null;
        }


    }
}
