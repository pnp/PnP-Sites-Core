using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectNavigation : ObjectHandlerBase
    {
        const string NavigationShowSiblings = "__NavigationShowSiblings";
        private bool ClearWarningShown = false;
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

                if (!WebSupportsExtractNavigation(web))
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Navigation_Context_web_is_not_publishing);
                    return template;
                }

                // Retrieve the current web navigation settings
                var navigationSettings = new WebNavigationSettings(web.Context, web);
                navigationSettings.EnsureProperties(ns => ns.AddNewPagesToNavigation, ns => ns.CreateFriendlyUrlsForNewPages,
                    ns => ns.CurrentNavigation, ns => ns.GlobalNavigation);

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

                var navigationEntity = new Model.Navigation(new GlobalNavigation(globalNavigationType,
                                                                globalNavigationType == GlobalNavigationType.Structural ? GetGlobalStructuralNavigation(web, navigationSettings) : null,
                                                                globalNavigationType == GlobalNavigationType.Managed ? GetGlobalManagedNavigation(web, navigationSettings) : null),
                                                            new CurrentNavigation(currentNavigationType,
                                                                currentNavigationType == CurrentNavigationType.Structural | currentNavigationType == CurrentNavigationType.StructuralLocal ? GetCurrentStructuralNavigation(web, navigationSettings) : null,
                                                                currentNavigationType == CurrentNavigationType.Managed ? GetCurrentManagedNavigation(web, navigationSettings) : null)
                                                            );

                navigationEntity.AddNewPagesToNavigation = navigationSettings.AddNewPagesToNavigation;
                navigationEntity.CreateFriendlyUrlsForNewPages = navigationSettings.CreateFriendlyUrlsForNewPages;

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    if (!navigationEntity.Equals(creationInfo.BaseTemplate.Navigation))
                    {
                        template.Navigation = navigationEntity;
                    }
                }
                else
                {
                    template.Navigation = navigationEntity;
                }
            }

            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Navigation != null)
                {
                    if (!WebSupportsProvisionNavigation(web, template))
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Navigation_Context_web_is_not_publishing);
                        return parser;
                    }

                    // Check if this is not a noscript site as navigation features are not supported
                    bool isNoScriptSite = web.IsNoScriptSite();

                    // Retrieve the current web navigation settings
                    var navigationSettings = new WebNavigationSettings(web.Context, web);
                    web.Context.Load(navigationSettings, ns => ns.CurrentNavigation, ns => ns.GlobalNavigation);
                    web.Context.ExecuteQueryRetry();

                    navigationSettings.AddNewPagesToNavigation = template.Navigation.AddNewPagesToNavigation;
                    navigationSettings.CreateFriendlyUrlsForNewPages = template.Navigation.CreateFriendlyUrlsForNewPages;

                    if (template.Navigation.GlobalNavigation != null)
                    {
                        switch (template.Navigation.GlobalNavigation.NavigationType)
                        {
                            case GlobalNavigationType.Inherit:
                                navigationSettings.GlobalNavigation.Source = StandardNavigationSource.InheritFromParentWeb;
                                web.Navigation.UseShared = true;
                                break;
                            case GlobalNavigationType.Managed:
                                if (template.Navigation.GlobalNavigation.ManagedNavigation == null)
                                {
                                    throw new ApplicationException(CoreResources.Provisioning_ObjectHandlers_Navigation_missing_global_managed_navigation);
                                }
                                navigationSettings.GlobalNavigation.Source = StandardNavigationSource.TaxonomyProvider;
                                navigationSettings.GlobalNavigation.TermStoreId = Guid.Parse(parser.ParseString(template.Navigation.GlobalNavigation.ManagedNavigation.TermStoreId));
                                navigationSettings.GlobalNavigation.TermSetId = Guid.Parse(parser.ParseString(template.Navigation.GlobalNavigation.ManagedNavigation.TermSetId));
                                web.Navigation.UseShared = false;
                                break;
                            case GlobalNavigationType.Structural:
                            default:
                                if (template.Navigation.GlobalNavigation.StructuralNavigation == null)
                                {
                                    throw new ApplicationException(CoreResources.Provisioning_ObjectHandlers_Navigation_missing_global_structural_navigation);
                                }
                                navigationSettings.GlobalNavigation.Source = StandardNavigationSource.PortalProvider;
                                web.Navigation.UseShared = false;

                                ProvisionGlobalStructuralNavigation(web,
                                    template.Navigation.GlobalNavigation.StructuralNavigation, parser, applyingInformation.ClearNavigation, scope);

                                break;
                        }
                        if (!isNoScriptSite)
                        {
                            navigationSettings.Update(TaxonomySession.GetTaxonomySession(web.Context));
                            web.Context.ExecuteQueryRetry();
                        }
                    }

                    if (template.Navigation.CurrentNavigation != null)
                    {
                        switch (template.Navigation.CurrentNavigation.NavigationType)
                        {
                            case CurrentNavigationType.Inherit:
                                navigationSettings.CurrentNavigation.Source = StandardNavigationSource.InheritFromParentWeb;
                                break;
                            case CurrentNavigationType.Managed:
                                if (template.Navigation.CurrentNavigation.ManagedNavigation == null)
                                {
                                    throw new ApplicationException(CoreResources.Provisioning_ObjectHandlers_Navigation_missing_current_managed_navigation);
                                }
                                navigationSettings.CurrentNavigation.Source = StandardNavigationSource.TaxonomyProvider;
                                navigationSettings.CurrentNavigation.TermStoreId = Guid.Parse(parser.ParseString(template.Navigation.CurrentNavigation.ManagedNavigation.TermStoreId));
                                navigationSettings.CurrentNavigation.TermSetId = Guid.Parse(parser.ParseString(template.Navigation.CurrentNavigation.ManagedNavigation.TermSetId));
                                break;
                            case CurrentNavigationType.StructuralLocal:
                                if (!isNoScriptSite)
                                {
                                    web.SetPropertyBagValue(NavigationShowSiblings, "false");
                                }
                                if (template.Navigation.CurrentNavigation.StructuralNavigation == null)
                                {
                                    throw new ApplicationException(CoreResources.Provisioning_ObjectHandlers_Navigation_missing_current_structural_navigation);
                                }
                                navigationSettings.CurrentNavigation.Source = StandardNavigationSource.PortalProvider;

                                ProvisionCurrentStructuralNavigation(web,
                                    template.Navigation.CurrentNavigation.StructuralNavigation, parser, applyingInformation.ClearNavigation, scope);

                                break;
                            case CurrentNavigationType.Structural:
                            default:
                                if (!isNoScriptSite)
                                {
                                    web.SetPropertyBagValue(NavigationShowSiblings, "true");
                                }
                                if (template.Navigation.CurrentNavigation.StructuralNavigation == null)
                                {
                                    throw new ApplicationException(CoreResources.Provisioning_ObjectHandlers_Navigation_missing_current_structural_navigation);
                                }
                                navigationSettings.CurrentNavigation.Source = StandardNavigationSource.PortalProvider;

                                ProvisionCurrentStructuralNavigation(web,
                                    template.Navigation.CurrentNavigation.StructuralNavigation, parser, applyingInformation.ClearNavigation, scope);

                                break;
                        }

                        if (!isNoScriptSite)
                        {
                            navigationSettings.Update(TaxonomySession.GetTaxonomySession(web.Context));
                            web.Context.ExecuteQueryRetry();
                        }
                    }
                }
            }

            return parser;
        }

        #region Utility methods

        private bool WebSupportsProvisionNavigation(Web web, ProvisioningTemplate template)
        {
            bool isNavSupported = true;
            // The Navigation handler for managed metedata only works for sites with Publishing Features enabled
            if (!web.IsPublishingWeb())
            {
                // NOTE: Here there could be a very edge case for a site where publishing features were enabled, 
                // configured managed navigation, and then disabled, keeping one navigation managed and another
                // one structural. Just as a reminder ...
                if (template.Navigation.GlobalNavigation != null
                    && template.Navigation.GlobalNavigation.NavigationType == GlobalNavigationType.Managed)
                {
                    isNavSupported = false;
                }
                if (template.Navigation.CurrentNavigation != null
                    && template.Navigation.CurrentNavigation.NavigationType == CurrentNavigationType.Managed)
                {
                    isNavSupported = false;
                }
            }
            return isNavSupported;
        }

        private bool WebSupportsExtractNavigation(Web web)
        {
            bool isNavSupported = true;
            // The Navigation handler for managed metedata only works for sites with Publishing Features enabled
            if (!web.IsPublishingWeb())
            {
                // NOTE: Here we could have the same edge case of method WebSupportsProvisionNavigation. 
                // Just as a reminder ...
                var navigationSettings = new WebNavigationSettings(web.Context, web);
                navigationSettings.EnsureProperties(ns => ns.CurrentNavigation, ns => ns.GlobalNavigation);
                if (navigationSettings.CurrentNavigation.Source == StandardNavigationSource.TaxonomyProvider)
                {
                    isNavSupported = false;
                }
                if (navigationSettings.GlobalNavigation.Source == StandardNavigationSource.TaxonomyProvider)
                {
                    isNavSupported = false;
                }
            }
            return isNavSupported;
        }

        private Boolean AreSiblingsEnabledForCurrentStructuralNavigation(Web web)
        {
            bool siblingsEnabled = false;

            if (bool.TryParse(web.GetPropertyBagValueString(NavigationShowSiblings, "false"), out siblingsEnabled))
            {
            }

            return siblingsEnabled;
        }

        private void ProvisionGlobalStructuralNavigation(Web web, StructuralNavigation structuralNavigation, TokenParser parser, bool clearNavigation, PnPMonitoredScope scope)
        {
            ProvisionStructuralNavigation(web, structuralNavigation, parser, false, clearNavigation, scope);
        }

        private void ProvisionCurrentStructuralNavigation(Web web, StructuralNavigation structuralNavigation, TokenParser parser, bool clearNavigation, PnPMonitoredScope scope)
        {
            ProvisionStructuralNavigation(web, structuralNavigation, parser, true, clearNavigation, scope);
        }

        private void ProvisionStructuralNavigation(Web web, StructuralNavigation structuralNavigation, TokenParser parser, bool currentNavigation, bool clearNavigation, PnPMonitoredScope scope)
        {
            // Determine the target structural navigation
            var navigationType = currentNavigation ?
                Enums.NavigationType.QuickLaunch :
                Enums.NavigationType.TopNavigationBar;
            if (structuralNavigation != null)
            {
                // Remove existing nodes, if requested
                if (structuralNavigation.RemoveExistingNodes || clearNavigation)
                {
                    if (!structuralNavigation.RemoveExistingNodes && !ClearWarningShown)
                    {
                        WriteMessage("You chose to override the template value RemoveExistingNodes=\"false\" by specifying ClearNavigation", ProvisioningMessageType.Warning);
                        ClearWarningShown = true;
                    }
                    web.DeleteAllNavigationNodes(navigationType);
                }

                // Provision root level nodes, and children recursively
                if (structuralNavigation.NavigationNodes.Any())
                {
                    ProvisionStructuralNavigationNodes(
                        web,
                        parser,
                        navigationType,
                        structuralNavigation.NavigationNodes,
                        scope
                    );
                }
            }
        }

        private void ProvisionStructuralNavigationNodes(Web web, TokenParser parser, Enums.NavigationType navigationType, Model.NavigationNodeCollection nodes, PnPMonitoredScope scope, string parentNodeTitle = null)
        {
            foreach (var node in nodes)
            {
                try
                {
                    var navNode = web.AddNavigationNode(
                        parser.ParseString(node.Title),
                        new Uri(parser.ParseString(node.Url), UriKind.RelativeOrAbsolute),
                        parser.ParseString(parentNodeTitle),
                        navigationType,
                        node.IsExternal
                        );

#if !SP2013
                    if (node.Title.ContainsResourceToken())
                    {
                        navNode.LocalizeNavigationNode(web, node.Title, parser, scope);
                    }
#endif
                }
                catch (ServerException ex)
                {
                    // If the SharePoint link doesn't exist, provision it as external link
                    // when we provision as external link, the server side URL validation won't kick-in
                    // This handles the "no such file or url found" error

                    WriteMessage(String.Format(CoreResources.Provisioning_ObjectHandlers_Navigation_Link_Provisioning_Failed_Retry, node.Title), ProvisioningMessageType.Warning);

                    if (ex.ServerErrorCode == -2130247147)
                    {
                        try
                        {
                            var navNode = web.AddNavigationNode(
                                parser.ParseString(node.Title),
                                new Uri(parser.ParseString(node.Url), UriKind.RelativeOrAbsolute),
                                parser.ParseString(parentNodeTitle),
                                navigationType,
                                true
                                );
                        }
                        catch (Exception innerEx)
                        {
                            WriteMessage(String.Format(CoreResources.Provisioning_ObjectHandlers_Navigation_Link_Provisioning_Failed, innerEx.Message), ProvisioningMessageType.Warning);
                        }
                    }
                    else
                    {
                        WriteMessage(String.Format(CoreResources.Provisioning_ObjectHandlers_Navigation_Link_Provisioning_Failed, ex.Message), ProvisioningMessageType.Warning);
                    }
                }

                ProvisionStructuralNavigationNodes(
                    web,
                    parser,
                    navigationType,
                    node.NavigationNodes,
                    scope,
                    parser.ParseString(node.Title));
            }
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

            // Apply any token replacement for taxonomy IDs
            TokenizeManagedNavigationTaxonomyIds(web, result);

            return (result);
        }

        private StructuralNavigation GetStructuralNavigation(Web web, WebNavigationSettings navigationSettings, Boolean currentNavigation)
        {
            // By default avoid removing existing nodes
            var result = new StructuralNavigation { RemoveExistingNodes = false };
            Microsoft.SharePoint.Client.NavigationNodeCollection sourceNodes = currentNavigation ?
                web.Navigation.QuickLaunch : web.Navigation.TopNavigationBar;

            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.Load(sourceNodes);
            web.Context.ExecuteQueryRetry();

            if (!sourceNodes.ServerObjectIsNull.Value)
            {
                result.NavigationNodes.AddRange(from n in sourceNodes.AsEnumerable()
                                                select n.ToDomainModelNavigationNode(web));
            }
            return (result);
        }

        protected void TokenizeManagedNavigationTaxonomyIds(Web web, ManagedNavigation managedNavigation)
        {
            // Replace Taxonomy field references to SspId, TermSetId with tokens
            TaxonomySession session = TaxonomySession.GetTaxonomySession(web.Context);
            TermStore defaultStore = session.GetDefaultSiteCollectionTermStore();
            var site = (web.Context as ClientContext).Site;
            var siteCollectionTermGroup = defaultStore.GetSiteCollectionGroup(site, false);
            web.Context.Load(siteCollectionTermGroup, t => t.Name);
            web.Context.ExecuteQueryRetry();
            string siteCollectionTermGroupName = null;
            if (!siteCollectionTermGroup.ServerObjectIsNull.Value)
            {
                siteCollectionTermGroupName = siteCollectionTermGroup.Name;
            }
            web.Context.Load(defaultStore, ts => ts.Name, ts => ts.Id);
            web.Context.ExecuteQueryRetry();

            Guid navigationTermStoreId = Guid.Parse(managedNavigation.TermStoreId);
            if (navigationTermStoreId != Guid.Empty)
            {
                TermStore navigationTermStore = session.TermStores.GetById(navigationTermStoreId);
                web.Context.Load(navigationTermStore, ts => ts.Name, ts => ts.Id);
                web.Context.ExecuteQueryRetry();

                if (!navigationTermStore.ServerObjectIsNull())
                {
                    if (navigationTermStore.Id == defaultStore.Id)
                    {
                        managedNavigation.TermStoreId = "{sitecollectiontermstoreid}";
                    }
                    else
                    {
                        managedNavigation.TermStoreId = $"{{termstoreid:{navigationTermStore.Name}}}";
                    }

                    Guid navigationTermSetId = Guid.Parse(managedNavigation.TermSetId);
                    if (navigationTermSetId != Guid.Empty)
                    {
                        var navigationTermSet = navigationTermStore.GetTermSet(navigationTermSetId);
                        web.Context.Load(navigationTermSet, ts => ts.Name, ts => ts.Id, ts => ts.Group);
                        web.Context.ExecuteQueryRetry();

                        if (!navigationTermSet.ServerObjectIsNull())
                        {
                            if (navigationTermSet.Group.Name == siteCollectionTermGroupName)
                            {
                                managedNavigation.TermSetId = $"{{sitecollectiontermsetid:{navigationTermSet.Name}}}";
                            }
                            else
                            {
                                managedNavigation.TermSetId =
                                    $"{{termsetid:{navigationTermSet.Group.Name}:{navigationTermSet.Name}}}";
                            }
                        }
                    }
                }
            }
        }

        #endregion

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return WebSupportsExtractNavigation(web);
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return (template.Navigation != null &&
                WebSupportsProvisionNavigation(web, template));
        }
    }

    internal static class NavigationNodeExtensions
    {
        internal static Model.NavigationNode ToDomainModelNavigationNode(this Microsoft.SharePoint.Client.NavigationNode node, Web web)
        {

            var result = new Model.NavigationNode
            {
                Title = node.Title,
                IsExternal = node.IsExternal,
                Url = web.ServerRelativeUrl != "/" ? node.Url.Replace(web.ServerRelativeUrl, "{site}") : $"{{site}}{node.Url}"
            };

            node.Context.Load(node.Children);
            node.Context.ExecuteQueryRetry();

            result.NavigationNodes.AddRange(from n in node.Children.AsEnumerable()
                                            select n.ToDomainModelNavigationNode(web));

            return (result);
        }
    }
}
