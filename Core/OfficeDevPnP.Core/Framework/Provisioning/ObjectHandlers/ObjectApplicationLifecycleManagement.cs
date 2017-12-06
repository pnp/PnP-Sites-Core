using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.ALM;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectApplicationLifecycleManagement : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Application Lifecycle Management"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // The ALM API do not support the local Site Collection App Catalog
                // Thus, so far we skip the AppCatalog section
                // NOOP

                // Process the collection of Apps installed in the current Site Collection
                var manager = new AppManager(web.Context as ClientContext);

                var siteApps = manager.GetAvailable()?.Where(a => a.InstalledVersion != null)?.ToList();
                if (siteApps != null && siteApps.Count > 0)
                {
                    foreach (var app in siteApps)
                    {
                        template.ApplicationLifecycleManagement.Apps.Add(new Model.App
                        {
                            AppId = app.Id.ToString(),
                            Action = AppAction.Install,
                        });
                    }
                }
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.ApplicationLifecycleManagement != null)
                {
                    var manager = new AppManager(web.Context as ClientContext);

                    // The ALM API do not support the local Site Collection App Catalog
                    // Thus, so far we skip the AppCatalog section
                    // NOOP

                    if (template.ApplicationLifecycleManagement.Apps != null &&
                        template.ApplicationLifecycleManagement.Apps.Count > 0)
                    {
                        foreach (var app in template.ApplicationLifecycleManagement.Apps)
                        {
                            var appId = Guid.Parse(parser.ParseString(app.AppId));

                            switch (app.Action)
                            {
                                case AppAction.Install:
                                    manager.Install(appId);
                                    break;
                                case AppAction.Uninstall:
                                    manager.Uninstall(appId);
                                    break;
                                case AppAction.Update:
                                    manager.Upgrade(appId);
                                    break;
                            }
                        }
                    }
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return (!web.IsSubSite());
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return (!web.IsSubSite() && template.ApplicationLifecycleManagement != null);
        }
    }
}
