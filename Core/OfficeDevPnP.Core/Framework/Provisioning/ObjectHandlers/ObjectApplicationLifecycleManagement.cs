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
#if !ONPREMISES
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
                var appCatalogUri = web.GetAppCatalog();
                if(appCatalogUri != null)
                {
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
                        //Get tenant app catalog
                        var appCatalogUri = web.GetAppCatalog();
                        if (appCatalogUri != null)
                        {
                            // Get the apps already installed in the site
                            var siteApps = manager.GetAvailable()?.Where(a => a.InstalledVersion != null)?.ToList();

                            foreach (var app in template.ApplicationLifecycleManagement.Apps)
                            {
                                var appId = Guid.Parse(parser.ParseString(app.AppId));
                                var alreadyExists = siteApps.Any(a => a.Id == appId);
                                var working = false;

                                if (app.Action == AppAction.Install && !alreadyExists)
                                {
                                    manager.Install(appId);
                                    working = true;
                                }
                                else if (app.Action == AppAction.Install && alreadyExists)
                                {
                                    WriteMessage($"App with ID {appId} already exists in the target site and will be skipped", ProvisioningMessageType.Warning);
                                }
                                else if (app.Action == AppAction.Uninstall && alreadyExists)
                                {
                                    manager.Uninstall(appId);
                                    working = true;
                                }
                                else if (app.Action == AppAction.Uninstall && !alreadyExists)
                                {
                                    WriteMessage($"App with ID {appId} does not exist in the target site and cannot be uninstalled", ProvisioningMessageType.Warning);
                                }
                                else if (app.Action == AppAction.Update && alreadyExists)
                                {
                                    manager.Upgrade(appId);
                                    working = true;
                                }
                                else if (app.Action == AppAction.Update && !alreadyExists)
                                {
                                    WriteMessage($"App with ID {appId} does not exist in the target site and cannot be updated", ProvisioningMessageType.Warning);
                                }

                                if (app.SyncMode == SyncMode.Synchronously && working)
                                {
                                    // We need to wait for the app management
                                    // to be completed before proceeding
                                }
                            }
                        }
                        else
                        {
                            WriteMessage($"Tenant app catalog doesn't exist. ALM step will be skipped.", ProvisioningMessageType.Warning);
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

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return (!web.IsSubSite() && template.ApplicationLifecycleManagement != null);
        }
    }
#endif
}
