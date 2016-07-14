using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using Microsoft.SharePoint.Client.Publishing.Navigation;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectNavigation : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Navigation"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                GlobalNavigationType globalNavigationType;
                CurrentNavigationType currentNavigationType;

                // The Navigation handler works only for sites with Publishing Features enabled
                if (!web.IsPublishingWeb())
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Navigation_Context_web_is_not_publishing);
                    return template;
                }

                // Retrieve the current web navigation settings
                var navigationSettings = new WebNavigationSettings(web.Context, web);
                web.Context.Load(navigationSettings, ns => ns.CurrentNavigation, ns => ns.GlobalNavigation);
                web.Context.ExecuteQueryRetry();

                switch (navigationSettings.GlobalNavigation.Source)
                {
                    case StandardNavigationSource.InheritFromParentWeb:
                        // Global Navigation is Inherited
                        globalNavigationType = GlobalNavigationType.Inherit;
                        break;
                    case StandardNavigationSource.TaxonomyProvider:
                        // Global Navigation is Managed
                        globalNavigationType = GlobalNavigationType.Managed;
                        break;
                    case StandardNavigationSource.PortalProvider:
                    default:
                        // Global Navigation is Structural
                        globalNavigationType = GlobalNavigationType.Structural;
                        break;
                }

                switch (navigationSettings.CurrentNavigation.Source)
                {
                    case StandardNavigationSource.InheritFromParentWeb:
                        // Current Navigation is Inherited
                        currentNavigationType = CurrentNavigationType.Inherit;
                        break;
                    case StandardNavigationSource.TaxonomyProvider:
                        // Current Navigation is Managed
                        currentNavigationType = CurrentNavigationType.Managed;
                        break;
                    case StandardNavigationSource.PortalProvider:
                    default:
                        // Current Navigation is Structural
                        if (AreSiblingsEnabledForCurrentStructuralNavigation(web))
                        {
                            currentNavigationType = CurrentNavigationType.Structural;
                        }
                        else
                        {
                            currentNavigationType = CurrentNavigationType.StructuralLocal;
                        }
                        break;
                }

                template.Navigation = new Model.Navigation(
                    new GlobalNavigation(globalNavigationType,
                        globalNavigationType == GlobalNavigationType.Structural ? GetGlobalStructuralNavigation(web, navigationSettings) : null,
                        globalNavigationType == GlobalNavigationType.Managed ? GetGlobalManagedNavigation(web, navigationSettings) : null),
                    new CurrentNavigation(currentNavigationType,
                        currentNavigationType == CurrentNavigationType.Structural | currentNavigationType == CurrentNavigationType.StructuralLocal ? GetCurrentStructuralNavigation(web, navigationSettings) : null,
                        currentNavigationType == CurrentNavigationType.Managed ? GetCurrentManagedNavigation(web, navigationSettings) : null)
                    );
            }

            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.AuditSettings != null)
                {
                    var site = (web.Context as ClientContext).Site;

                    site.EnsureProperties(s => s.Audit, s => s.AuditLogTrimmingRetention, s => s.TrimAuditLog);

                    var siteAuditSettings = site.Audit;

                    var isDirty = false;
                    if (template.AuditSettings.AuditFlags != siteAuditSettings.AuditFlags)
                    {
                        site.Audit.AuditFlags = template.AuditSettings.AuditFlags;
                        site.Audit.Update();
                        isDirty = true;
                    }
                    if (template.AuditSettings.AuditLogTrimmingRetention != site.AuditLogTrimmingRetention)
                    {
                        site.AuditLogTrimmingRetention = template.AuditSettings.AuditLogTrimmingRetention;
                        isDirty = true;
                    }
                    if (template.AuditSettings.TrimAuditLog != site.TrimAuditLog)
                    {
                        site.TrimAuditLog = template.AuditSettings.TrimAuditLog;
                        isDirty = true;
                    }
                    if (isDirty)
                    {
                        web.Context.ExecuteQueryRetry();
                    }
                }
            }

            return parser;
        }

        #region Utility methods

        private Boolean AreSiblingsEnabledForCurrentStructuralNavigation(Web web)
        {
            bool siblingsEnabled = false;

            if (bool.TryParse(web.GetPropertyBagValueString("__NavigationShowSiblings", "false"), out siblingsEnabled))
            {
            }

            return siblingsEnabled;
        }

        private ManagedNavigation GetGlobalManagedNavigation(Web web, WebNavigationSettings navigationSettings)
        {
            return GetManagedNavigation(web, navigationSettings, false);
        }

        private StructuralNavigation GetGlobalStructuralNavigation(Web web, WebNavigationSettings navigationSettings)
        {
            return GetStructuralNavigation(web, navigationSettings, false);
        }

        private ManagedNavigation GetCurrentManagedNavigation(Web web, WebNavigationSettings navigationSettings)
        {
            return GetManagedNavigation(web, navigationSettings, true);
        }

        private StructuralNavigation GetCurrentStructuralNavigation(Web web, WebNavigationSettings navigationSettings)
        {
            return GetStructuralNavigation(web, navigationSettings, true);
        }

        private ManagedNavigation GetManagedNavigation(Web web, WebNavigationSettings navigationSettings, Boolean currentNavigation)
        {
            var result = new ManagedNavigation
            {
                TermStoreId = currentNavigation ? navigationSettings.CurrentNavigation.TermStoreId.ToString() : navigationSettings.GlobalNavigation.TermStoreId.ToString(),
                TermSetId = currentNavigation ? navigationSettings.CurrentNavigation.TermSetId.ToString() : navigationSettings.GlobalNavigation.TermSetId.ToString(),
            };
            return (result);
        }

        private StructuralNavigation GetStructuralNavigation(Web web, WebNavigationSettings navigationSettings, Boolean currentNavigation)
        {
            // By default avoid removing existing nodes
            var result = new StructuralNavigation { RemoveExistingNodes = false };
            Microsoft.SharePoint.Client.NavigationNodeCollection sourceNodes = currentNavigation ?
                web.Navigation.QuickLaunch : web.Navigation.TopNavigationBar;

            result.NavigationNodes.AddRange(from n in sourceNodes
                                            select n.ToDomainModelNavigationNode());

            return (result);
        }

        #endregion

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return web.IsPublishingWeb();
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return web.IsPublishingWeb() && template.Navigation != null;
        }
    }

    internal static class NavigationNodeExtensions
    {
        internal static Model.NavigationNode ToDomainModelNavigationNode(this Microsoft.SharePoint.Client.NavigationNode node)
        {
            var result = new Model.NavigationNode
            {
                Title = node.Title,
                IsExternal = node.IsExternal,
                Url = node.Url,
            };

            result.NavigationNodes.AddRange(from n in node.Children
                                            select n.ToDomainModelNavigationNode());

            return (result);
        }
    }
}
