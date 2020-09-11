using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Enums;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Collections;
using System.Linq.Expressions;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Utilities;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using Newtonsoft.Json.Linq;

namespace Microsoft.SharePoint.Client
{

    /// <summary>
    /// This class holds navigation related methods
    /// </summary>
    public static partial class NavigationExtensions
    {

        #region Area Navigation (publishing sites)
        const string PublishingFeatureActivated = "__PublishingFeatureActivated";
        const string WebNavigationSettings = "_webnavigationsettings";
        const string CurrentNavigationIncludeTypes = "__CurrentNavigationIncludeTypes";
        const string CurrentDynamicChildLimit = "__CurrentDynamicChildLimit";
        const string GlobalNavigationIncludeTypes = "__GlobalNavigationIncludeTypes";
        const string GlobalDynamicChildLimit = "__GlobalDynamicChildLimit";
        const string NavigationOrderingMethod = "__NavigationOrderingMethod";
        const string NavigationAutomaticSortingMethod = "__NavigationAutomaticSortingMethod";
        const string NavigationSortAscending = "__NavigationSortAscending";
        const string NavigationShowSiblings = "__NavigationShowSiblings";

        /// <summary>
        /// Returns the navigation settings for the selected web
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <returns>Returns AreaNavigationEntity settings</returns>
        public static AreaNavigationEntity GetNavigationSettings(this Web web)
        {
            var nav = new AreaNavigationEntity();

            //Read all the properties of the web
            web.Context.Load(web, w => w.AllProperties);
            web.Context.ExecuteQueryRetry();

            // Determine if managed navigation is used...if so the other properties are not relevant
            string webNavigationSettings = web.AllProperties.GetPropertyAsString(WebNavigationSettings);
            if (webNavigationSettings == null)
            {
                nav.CurrentNavigation.ManagedNavigation = false;
                nav.GlobalNavigation.ManagedNavigation = false;
            }
            else
            {
                var navigationSettings = XElement.Parse(webNavigationSettings);
                IEnumerable<XElement> navNodes = navigationSettings.XPathSelectElements("./SiteMapProviderSettings/TaxonomySiteMapProviderSettings");
                foreach (var node in navNodes)
                {
                    if (node.Attribute("Name").Value.Equals("CurrentNavigationTaxonomyProvider", StringComparison.InvariantCulture))
                    {
                        bool managedNavigation = true;
                        if (node.Attribute("Disabled") != null)
                        {
                            if (bool.TryParse(node.Attribute("Disabled").Value, out managedNavigation))
                            {
                                managedNavigation = false;
                            }
                        }
                        nav.CurrentNavigation.ManagedNavigation = managedNavigation;
                    }
                    else if (node.Attribute("Name").Value.Equals("GlobalNavigationTaxonomyProvider", StringComparison.InvariantCulture))
                    {
                        var managedNavigation = true;
                        if (node.Attribute("Disabled") != null)
                        {
                            if (bool.TryParse(node.Attribute("Disabled").Value, out managedNavigation))
                            {
                                managedNavigation = false;
                            }
                        }
                        nav.GlobalNavigation.ManagedNavigation = managedNavigation;
                    }
                }

                // Get settings related to page creation
                XElement pageNode = navigationSettings.XPathSelectElement("./NewPageSettings");
                if (pageNode != null)
                {
                    if (pageNode.Attribute("AddNewPagesToNavigation") != null)
                    {
                        bool addNewPagesToNavigation;
                        if (bool.TryParse(pageNode.Attribute("AddNewPagesToNavigation").Value, out addNewPagesToNavigation))
                        {
                            nav.AddNewPagesToNavigation = addNewPagesToNavigation;
                        }
                    }

                    if (pageNode.Attribute("CreateFriendlyUrlsForNewPages") != null)
                    {
                        bool createFriendlyUrlsForNewPages;
                        if (bool.TryParse(pageNode.Attribute("CreateFriendlyUrlsForNewPages").Value, out createFriendlyUrlsForNewPages))
                        {
                            nav.CreateFriendlyUrlsForNewPages = createFriendlyUrlsForNewPages;
                        }
                    }
                }

                // Get navigation inheritance
                IEnumerable<XElement> switchableNavNodes = navigationSettings.XPathSelectElements("./SiteMapProviderSettings/SwitchableSiteMapProviderSettings");
                foreach (var node in switchableNavNodes)
                {
                    if (node.Attribute("Name").Value.Equals("CurrentNavigationSwitchableProvider", StringComparison.InvariantCulture))
                    {
                        bool inherit = false;
                        if (node.Attribute("UseParentSiteMap") != null)
                        {
                            bool.TryParse(node.Attribute("UseParentSiteMap").Value, out inherit);
                        }
                        nav.CurrentNavigation.InheritFromParentWeb = inherit;
                    }
                    else if (node.Attribute("Name").Value.Equals("GlobalNavigationSwitchableProvider", StringComparison.InvariantCulture))
                    {
                        bool inherit = false;
                        if (node.Attribute("UseParentSiteMap") != null)
                        {
                            bool.TryParse(node.Attribute("UseParentSiteMap").Value, out inherit);
                        }
                        nav.GlobalNavigation.InheritFromParentWeb = inherit;
                    }
                }
            }

            // Only read the other values that make sense when not using managed navigation
            if (!nav.CurrentNavigation.ManagedNavigation)
            {
                var currentNavigationIncludeTypes = web.AllProperties.GetPropertyAsInt(CurrentNavigationIncludeTypes);
                if (currentNavigationIncludeTypes > -1)
                {
                    MapFromNavigationIncludeTypes(nav.CurrentNavigation, currentNavigationIncludeTypes);
                }

                var currentDynamicChildLimit = web.AllProperties.GetPropertyAsInt(CurrentDynamicChildLimit);
                if (currentDynamicChildLimit > -1)
                {
                    nav.CurrentNavigation.MaxDynamicItems = currentDynamicChildLimit;
                }

                // For the current navigation there's an option to show the sites siblings in structural navigation
                if (web.IsSubSite())
                {
                    var showSiblings = false;
                    var navigationShowSiblings = web.AllProperties.GetPropertyAsString(NavigationShowSiblings);
                    if (bool.TryParse(navigationShowSiblings, out showSiblings))
                    {
                        nav.CurrentNavigation.ShowSiblings = showSiblings;
                    }
                }
            }

            if (!nav.GlobalNavigation.ManagedNavigation)
            {
                var globalNavigationIncludeTypes = web.AllProperties.GetPropertyAsInt(GlobalNavigationIncludeTypes);
                if (globalNavigationIncludeTypes > -1)
                {
                    MapFromNavigationIncludeTypes(nav.GlobalNavigation, globalNavigationIncludeTypes);
                }

                var globalDynamicChildLimit = web.AllProperties.GetPropertyAsInt(GlobalDynamicChildLimit);
                if (globalDynamicChildLimit > -1)
                {
                    nav.GlobalNavigation.MaxDynamicItems = globalDynamicChildLimit;
                }
            }

            // Read the sorting value 
            var navigationOrderingMethod = web.AllProperties.GetPropertyAsInt(NavigationOrderingMethod);
            if (navigationOrderingMethod > -1)
            {
                nav.Sorting = (StructuralNavigationSorting)navigationOrderingMethod;
            }

            // Read the sort by value
            var navigationAutomaticSortingMethod = web.AllProperties.GetPropertyAsInt(NavigationAutomaticSortingMethod);
            if (navigationAutomaticSortingMethod > -1)
            {
                nav.SortBy = (StructuralNavigationSortBy)navigationAutomaticSortingMethod;
            }

            // Read the ordering setting
            var navigationSortAscending = true;
            var navProp = web.AllProperties.GetPropertyAsString(NavigationSortAscending);

            if (bool.TryParse(navProp, out navigationSortAscending))
            {
                nav.SortAscending = navigationSortAscending;
            }

            return nav;
        }

        /// <summary>
        /// Updates navigation settings for the current web
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="navigationSettings">Navigation settings to update</param>
        public static void UpdateNavigationSettings(this Web web, AreaNavigationEntity navigationSettings)
        {
            //Read all the properties of the web
            web.Context.Load(web, w => w.AllProperties);
            web.Context.ExecuteQueryRetry();

            if (!ArePublishingFeaturesActivated(web.AllProperties))
            {
                throw new ArgumentException("Structural navigation settings are only supported for publishing sites");
            }

            // Use publishing CSOM API to switch between managed metadata and structural navigation
            var taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
            web.Context.Load(taxonomySession);
            web.Context.ExecuteQueryRetry();
            var webNav = new WebNavigationSettings(web.Context, web);
            if (navigationSettings.GlobalNavigation.InheritFromParentWeb)
            {
                if (web.IsSubSite())
                {
                    webNav.GlobalNavigation.Source = StandardNavigationSource.InheritFromParentWeb;
                }
                else
                {
                    throw new ArgumentException("Cannot inherit global navigation on root site.");
                }
            }
            else if (!navigationSettings.GlobalNavigation.ManagedNavigation)
            {
                webNav.GlobalNavigation.Source = StandardNavigationSource.PortalProvider;
            }
            else
            {
                webNav.GlobalNavigation.Source = StandardNavigationSource.TaxonomyProvider;
            }

            if (navigationSettings.CurrentNavigation.InheritFromParentWeb)
            {
                if (web.IsSubSite())
                {
                    webNav.CurrentNavigation.Source = StandardNavigationSource.InheritFromParentWeb;
                }
                else
                {
                    throw new ArgumentException("Cannot inherit current navigation on root site.");
                }
            }
            else if (!navigationSettings.CurrentNavigation.ManagedNavigation)
            {
                webNav.CurrentNavigation.Source = StandardNavigationSource.PortalProvider;
            }
            else
            {
                webNav.CurrentNavigation.Source = StandardNavigationSource.TaxonomyProvider;
            }

            // If managed metadata navigation is used, set settings related to page creation
            if (navigationSettings.GlobalNavigation.ManagedNavigation || navigationSettings.CurrentNavigation.ManagedNavigation)
            {
                webNav.AddNewPagesToNavigation = navigationSettings.AddNewPagesToNavigation;
                webNav.CreateFriendlyUrlsForNewPages = navigationSettings.CreateFriendlyUrlsForNewPages;
            }

            webNav.Update(taxonomySession);
            web.Context.ExecuteQueryRetry();

            //Read all the properties of the web again after the above update
            web.Context.Load(web, w => w.AllProperties);
            web.Context.ExecuteQueryRetry();

            if (!navigationSettings.GlobalNavigation.ManagedNavigation)
            {
                var globalNavigationIncludeType = MapToNavigationIncludeTypes(navigationSettings.GlobalNavigation);
                web.AllProperties[GlobalNavigationIncludeTypes] = globalNavigationIncludeType;
                web.AllProperties[GlobalDynamicChildLimit] = navigationSettings.GlobalNavigation.MaxDynamicItems;
            }

            if (!navigationSettings.CurrentNavigation.ManagedNavigation)
            {
                var currentNavigationIncludeType = MapToNavigationIncludeTypes(navigationSettings.CurrentNavigation);
                web.AllProperties[CurrentNavigationIncludeTypes] = currentNavigationIncludeType;
                web.AllProperties[CurrentDynamicChildLimit] = navigationSettings.CurrentNavigation.MaxDynamicItems;

                // Call web.update before the IsSubSite call as this might do an ExecuteQuery. Without the update called the changes will be lost
                web.Update();
                // For the current navigation there's an option to show the sites siblings in structural navigation
                if (web.IsSubSite())
                {
                    web.AllProperties[NavigationShowSiblings] = navigationSettings.CurrentNavigation.ShowSiblings.ToString();
                }
            }

            // if there's either global or current structural navigation then update the sorting settings
            if (!navigationSettings.GlobalNavigation.ManagedNavigation || !navigationSettings.CurrentNavigation.ManagedNavigation)
            {
                // If there's automatic sorting or pages are shown with automatic page sorting then we can set all sort options
                if ((navigationSettings.Sorting == StructuralNavigationSorting.Automatically) ||
                    (navigationSettings.Sorting == StructuralNavigationSorting.ManuallyButPagesAutomatically && (navigationSettings.GlobalNavigation.ShowPages || navigationSettings.CurrentNavigation.ShowPages)))
                {
                    // All sort options can be set
                    web.AllProperties[NavigationOrderingMethod] = (int)navigationSettings.Sorting;
                    web.AllProperties[NavigationAutomaticSortingMethod] = (int)navigationSettings.SortBy;
                    web.AllProperties[NavigationSortAscending] = navigationSettings.SortAscending.ToString();
                }
                else
                {
                    // if pages are not shown we can set sorting to either automatic or manual
                    if (!navigationSettings.GlobalNavigation.ShowPages && !navigationSettings.CurrentNavigation.ShowPages)
                    {
                        if (navigationSettings.Sorting == StructuralNavigationSorting.ManuallyButPagesAutomatically)
                        {
                            throw new ArgumentException("Sorting can only be set to StructuralNavigationSorting.ManuallyButPagesAutomatically when ShowPages has been selected in either the global or current structural navigation settings");
                        }
                    }

                    web.AllProperties[NavigationOrderingMethod] = (int)navigationSettings.Sorting;
                }
            }

            //Persist all property updates at once
            web.Update();
            web.Context.ExecuteQueryRetry();
        }

        private static int MapToNavigationIncludeTypes(StructuralNavigationEntity sne)
        {
            int navigationIncludeType = -1;

            if (!sne.ShowPages && !sne.ShowSubsites)
            {
                navigationIncludeType = 0;
            }
            else if (!sne.ShowPages && sne.ShowSubsites)
            {
                navigationIncludeType = 1;
            }
            else if (sne.ShowPages && !sne.ShowSubsites)
            {
                navigationIncludeType = 2;
            }
            else if (sne.ShowPages && sne.ShowSubsites)
            {
                navigationIncludeType = 3;
            }

            return navigationIncludeType;
        }


        private static void MapFromNavigationIncludeTypes(StructuralNavigationEntity sne, int navigationIncludeTypes)
        {
            if (navigationIncludeTypes == 0)
            {
                sne.ShowPages = false;
                sne.ShowSubsites = false;
            }
            else if (navigationIncludeTypes == 1)
            {
                sne.ShowPages = false;
                sne.ShowSubsites = true;
            }
            else if (navigationIncludeTypes == 2)
            {
                sne.ShowPages = true;
                sne.ShowSubsites = false;
            }
            else if (navigationIncludeTypes == 3)
            {
                sne.ShowPages = true;
                sne.ShowSubsites = true;
            }
        }

        private static bool ArePublishingFeaturesActivated(PropertyValues props)
        {
            var activated = false;

            if (bool.TryParse(props.GetPropertyAsString(PublishingFeatureActivated), out activated))
            {
            }

            return activated;
        }

        private static string GetPropertyAsString(this PropertyValues props, string key)
        {
            if (props.FieldValues.ContainsKey(key))
            {
                return props.FieldValues[key].ToString();
            }
            else
            {
                return null;
            }
        }
        private static int GetPropertyAsInt(this PropertyValues props, string key)
        {
            if (props.FieldValues.ContainsKey(key))
            {
                int res;
                if (int.TryParse(props.FieldValues[key].ToString(), out res))
                {
                    return res;
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                return -1;
            }
        }
        #endregion

        #region Managed Navigation (publishing sites)

        /// <summary>
        /// Returns an editable version of the Global Navigation TermSet for a web site
        /// </summary>
        /// <param name="web">The target web.</param>
        /// <param name="navigationKind">Declares whether to look for Current or Global Navigation</param>
        /// <returns>The editable Global Navigation TermSet</returns>
        public static NavigationTermSet GetEditableNavigationTermSet(this Web web, ManagedNavigationKind navigationKind)
        {
            if (!web.IsManagedNavigationEnabled(navigationKind))
            {
                throw new ApplicationException($"The current web is not using the Taxonomy provider for {navigationKind} Navigation.");
            }

            switch (navigationKind)
            {
                case ManagedNavigationKind.Global:
                    return (GetEditableNavigationTermSetByProviderName(web, web.Context,
                        "GlobalNavigationTaxonomyProvider"));
                case ManagedNavigationKind.Current:
                    return (GetEditableNavigationTermSetByProviderName(web, web.Context,
                        "CurrentNavigationTaxonomyProvider"));
                default:
                    return (null);
            }
        }

        /// <summary>
        /// Determines whether the current Web has the managed navigation enabled
        /// </summary>
        /// <param name="web">The target web.</param>
        /// <param name="navigationKind">The kind of navigation (Current or Global).</param>
        /// <returns>A boolean result of the test.</returns>
        public static bool IsManagedNavigationEnabled(this Web web, ManagedNavigationKind navigationKind)
        {
            var result = false;
            var navigationSettings = new WebNavigationSettings(web.Context, web);
            web.Context.Load(navigationSettings, ns => ns.CurrentNavigation, ns => ns.GlobalNavigation);
            web.Context.Load(web.ParentWeb, pw => pw.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            var targetNavigationSettings =
                navigationKind == ManagedNavigationKind.Current ?
                navigationSettings.CurrentNavigation : navigationSettings.GlobalNavigation;

            if (targetNavigationSettings.Source == StandardNavigationSource.InheritFromParentWeb &&
                !web.ParentWeb.ServerObjectIsNull())
            {
                var currentWebUri = new Uri(web.Url);
                var parentWebUri = new Uri($"{currentWebUri.Scheme}://{currentWebUri.Host}{web.ParentWeb.ServerRelativeUrl}");

                using (var parentContext = web.Context.Clone(parentWebUri))
                {
                    result = IsManagedNavigationEnabled(parentContext.Web, navigationKind);
                }
            }
            else
            {
                result = targetNavigationSettings.Source == StandardNavigationSource.TaxonomyProvider;
            }

            return (result);
        }

        private static NavigationTermSet GetEditableNavigationTermSetByProviderName(
            Web web, ClientRuntimeContext context, string providerName)
        {
            // Get the current taxonomy session and update cache, just in case
            var taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
            taxonomySession.UpdateCache();

            context.ExecuteQueryRetry();

            // Retrieve the Navigation TermSet for the current web
            var navigationTermSet = TaxonomyNavigation.GetTermSetForWeb(web.Context,
                web, providerName, true);
            context.Load(navigationTermSet);
            context.ExecuteQueryRetry();

            // Retrieve an editable TermSet for the current target navigation
            var editableNavigationTermSet = navigationTermSet.GetAsEditable(taxonomySession);
            context.Load(editableNavigationTermSet);
            context.ExecuteQueryRetry();

            return (editableNavigationTermSet);
        }

        #endregion

        #region Navigation elements - quicklaunch, top navigation, search navigation


        /// <summary>
        /// Add a node to quick launch, top navigation bar or search navigation. The node will be added as the last node in the
        /// collection.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="nodeTitle">the title of node to add</param>
        /// <param name="nodeUri">the URL of node to add</param>
        /// <param name="parentNodeTitle">if string.Empty, then will add this node as top level node. Contains the title of the immediate parent node, for third level nodes, providing <paramref name="l1ParentNodeTitle"/> is required.</param>
        /// <param name="navigationType">the type of navigation, quick launch, top navigation or search navigation</param>
        /// <param name="isExternal">true if the link is an external link</param>
        /// <param name="asLastNode">true if the link should be added as the last node of the collection</param>
        /// <param name="l1ParentNodeTitle">title of the first level parent, if this node is a third level navigation node</param>
        /// <returns>Newly added NavigationNode</returns>
        public static NavigationNode AddNavigationNode(this Web web, string nodeTitle, Uri nodeUri, string parentNodeTitle, NavigationType navigationType, bool isExternal = false, bool asLastNode = true, string l1ParentNodeTitle = null)
        {
            web.Context.Load(web, w => w.Navigation.QuickLaunch, w => w.Navigation.TopNavigationBar);
            web.Context.ExecuteQueryRetry();
            var node = new NavigationNodeCreationInformation
            {
                AsLastNode = asLastNode,
                Title = nodeTitle,
                Url = nodeUri != null ? nodeUri.OriginalString : string.Empty,
                IsExternal = isExternal
            };

            NavigationNode navigationNode = null;
            try
            {
                if (navigationType == NavigationType.QuickLaunch)
                {
                    var quickLaunch = web.Navigation.QuickLaunch;
                    if (string.IsNullOrEmpty(parentNodeTitle))
                    {
                        navigationNode = quickLaunch.Add(node);
                    }
                    else
                    {
                        navigationNode = CreateNodeAsChild(web, quickLaunch, node, parentNodeTitle, l1ParentNodeTitle);
                    }
                }
                else if (navigationType == NavigationType.TopNavigationBar)
                {
                    var topLink = web.Navigation.TopNavigationBar;
                    if (!string.IsNullOrEmpty(parentNodeTitle))
                    {
                        navigationNode = CreateNodeAsChild(web, topLink, node, parentNodeTitle, l1ParentNodeTitle);
                    }
                    else
                    {
                        navigationNode = topLink.Add(node);
                    }
                }
                else if (navigationType == NavigationType.SearchNav)
                {
                    var searchNavigation = web.LoadSearchNavigation();
                    navigationNode = searchNavigation.Add(node);
                }
            }
            finally
            {
                web.Context.ExecuteQueryRetry();
            }
            return navigationNode;
        }

        /// <summary>
        /// Creates a navigation node as a child of another (first or second level) navigation node.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="parentNodes">Level one nodes under which the node should be created</param>
        /// <param name="nodeToCreate">Node information</param>
        /// <param name="parentNodeTitle">The title of the immediate parent node (level two if child should be level three, level one otherwise)</param>
        /// <param name="l1ParentNodeTitle">The level one parent title or null, if the node to be created should be a level two node</param>
        /// <returns></returns>
        private static NavigationNode CreateNodeAsChild(Web web, NavigationNodeCollection parentNodes, NavigationNodeCreationInformation nodeToCreate, string parentNodeTitle, string l1ParentNodeTitle)
        {
            if (l1ParentNodeTitle != null)
            {
                var l1ParentNode = parentNodes.FirstOrDefault(n => n.Title.Equals(l1ParentNodeTitle, StringComparison.InvariantCultureIgnoreCase));
                if (l1ParentNode == null)
                {
                    return null;
                }
                web.Context.Load(l1ParentNode.Children);
                web.Context.ExecuteQueryRetry();
                parentNodes = l1ParentNode.Children;
            }

            var parentNode = parentNodes.FirstOrDefault(n => n.Title.Equals(parentNodeTitle, StringComparison.InvariantCultureIgnoreCase));
            var navigationNode = parentNode?.Children.Add(nodeToCreate);
            return navigationNode;
        }

        /// <summary>
        /// Deletes a navigation node from the quickLaunch or top navigation bar
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="nodeTitle">the title of node to delete</param>
        /// <param name="parentNodeTitle">if string.Empty, then will delete this node as top level node</param>
        /// <param name="navigationType">the type of navigation, quick launch, top navigation or search navigation</param>
        public static void DeleteNavigationNode(this Web web, string nodeTitle, string parentNodeTitle, NavigationType navigationType)
        {
            web.Context.Load(web, w => w.Navigation.QuickLaunch, w => w.Navigation.TopNavigationBar);
            web.Context.ExecuteQueryRetry();
            NavigationNode deleteNode = null;
            try
            {
                if (navigationType == NavigationType.QuickLaunch)
                {
                    var quickLaunch = web.Navigation.QuickLaunch;
                    if (string.IsNullOrEmpty(parentNodeTitle))
                    {
                        deleteNode = quickLaunch.SingleOrDefault(n => n.Title == nodeTitle);
                    }
                    else
                    {
                        foreach (var nodeInfo in quickLaunch)
                        {
                            if (nodeInfo.Title != parentNodeTitle)
                            {
                                continue;
                            }

                            web.Context.Load(nodeInfo.Children);
                            web.Context.ExecuteQueryRetry();
                            deleteNode = nodeInfo.Children.SingleOrDefault(n => n.Title == nodeTitle);
                        }
                    }
                }
                else if (navigationType == NavigationType.TopNavigationBar)
                {
                    var topLink = web.Navigation.TopNavigationBar;
                    if (string.IsNullOrEmpty(parentNodeTitle))
                    {
                        deleteNode = topLink.SingleOrDefault(n => n.Title == nodeTitle);
                    }
                    else
                    {
                        foreach (var nodeInfo in topLink)
                        {
                            if (nodeInfo.Title != parentNodeTitle)
                            {
                                continue;
                            }
                            web.Context.Load(nodeInfo.Children);
                            web.Context.ExecuteQueryRetry();
                            deleteNode = nodeInfo.Children.SingleOrDefault(n => n.Title == nodeTitle);
                        }
                    }
                }
                else if (navigationType == NavigationType.SearchNav)
                {
                    var nodeCollection = web.LoadSearchNavigation();
                    deleteNode = nodeCollection.SingleOrDefault(n => n.Title == nodeTitle);
                }
            }
            finally
            {
                deleteNode?.DeleteObject();
                web.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Deletes all Navigation Nodes from a given navigation
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="navigationType">The type of navigation to support</param>
        public static void DeleteAllNavigationNodes(this Web web, NavigationType navigationType)
        {
            if (navigationType == NavigationType.QuickLaunch)
            {
                web.Context.Load(web, w => w.Navigation.QuickLaunch);
                web.Context.ExecuteQueryRetry();

                var quickLaunch = web.Navigation.QuickLaunch;
                for (var i = quickLaunch.Count - 1; i >= 0; i--)
                {
                    quickLaunch[i].DeleteObject();
                }
                web.Context.ExecuteQueryRetry();
            }
            else if (navigationType == NavigationType.TopNavigationBar)
            {
                web.Context.Load(web, w => w.Navigation.TopNavigationBar);
                web.Context.ExecuteQueryRetry();
                var topNavigation = web.Navigation.TopNavigationBar;
                for (var i = topNavigation.Count - 1; i >= 0; i--)
                {
                    topNavigation[i].DeleteObject();
                }
                web.Context.ExecuteQueryRetry();
            }
            else if (navigationType == NavigationType.SearchNav)
            {
                var searchNavigation = web.LoadSearchNavigation();
                for (var i = searchNavigation.Count - 1; i >= 0; i--)
                {
                    searchNavigation[i].DeleteObject();
                }
                web.Context.ExecuteQueryRetry();
#if !ONPREMISES
            }
            else if (navigationType == NavigationType.Footer)
            {
                var footerNavigation = web.LoadFooterNavigation();
                if (footerNavigation != null)
                {
                    for (var i = footerNavigation.Count - 1; i >= 0; i--)
                    {
                        footerNavigation[i].DeleteObject();
                    }
                    web.Context.ExecuteQueryRetry();
                }
#endif
            }
        }

        /// <summary>
        /// Updates the navigation inheritance setting
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="inheritNavigation">boolean indicating if navigation inheritance is needed or not</param>
        public static void UpdateNavigationInheritance(this Web web, bool inheritNavigation)
        {
            web.Navigation.UseShared = inheritNavigation;
            web.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Loads the search navigation nodes
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <returns>Collection of NavigationNode instances</returns>
        public static NavigationNodeCollection LoadSearchNavigation(this Web web)
        {
            try
            {
                var searchNav = web.Navigation.GetNodeById(1040); // 1040 is the id of the search navigation            
                var nodeCollection = searchNav.Children;
                web.Context.Load(searchNav);
                web.Context.Load(nodeCollection);
                web.Context.ExecuteQueryRetry();
                return nodeCollection;
            }
            catch
            {
                 return null;
            }
        }

#if !ONPREMISES
        /// <summary>
        /// Returns the navigation elements shown in the footer
        /// </summary>
        /// <param name="web">Web instance to return the footer navigation of</param>
        /// <returns>NavigationNodeCollection containing the navigation elements shown in the footer or NULL if no navigation has been set on the footer</returns>
        public static NavigationNodeCollection LoadFooterNavigation(this Web web)
        {
            var structureString = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'").GetAwaiter().GetResult();
            var menuState = JObject.Parse(structureString);

            if (menuState["StartingNodeKey"] == null)
            {
                web.EnsureProperties(w => w.ServerRelativeUrl);
                var now = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss:Z");
                web.ExecutePostAsync($"/_api/navigation/SaveMenuState", $@"{{ ""menuState"":{{ ""Version"":""{now}"",""StartingNodeTitle"":""3a94b35f-030b-468e-80e3-b75ee84ae0ad"",""SPSitePrefix"":""/"",""SPWebPrefix"":""{web.ServerRelativeUrl}"",""FriendlyUrlPrefix"":"""",""SimpleUrl"":"""",""Nodes"":[]}}}}").GetAwaiter().GetResult();
                structureString = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'").GetAwaiter().GetResult();
                menuState = JObject.Parse(structureString);
            }

            if (menuState["Nodes"] != null)
            {
                var nodes = menuState["Nodes"] as JArray;
                var topNode = web.Navigation.GetNodeById(Convert.ToInt32(menuState["StartingNodeKey"].Value<string>()));
                web.Context.Load(topNode, n => n.Children.IncludeWithDefaultProperties());
                web.Context.ExecuteQueryRetry();
                var menuNode = topNode.Children.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_MENUNODEKEY);
                if (menuNode == null)
                {
                    // No navigation elements have been added to the footer
                    return null;
                }

                // Navigation elements have been added to the footer, return them
                menuNode.EnsureProperty(n => n.Children.IncludeWithDefaultProperties());
                return menuNode.Children;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the title shown in the footer
        /// </summary>
        /// <param name="web">Web instance to return the footer title of</param>
        /// <returns>Title shown in the footer or NULL if no title has been set</returns>
        public static string GetFooterTitle(this Web web)
        {
            var structureString = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'").GetAwaiter().GetResult();
            var menuState = JObject.Parse(structureString);

            if (menuState["Nodes"] == null)
            {
                // No information is returned which helps us to identity the title node
                return null;
            }

            // Retrieve the Node representing the title node
            var titleNode = menuState["Nodes"].FirstOrDefault(n => n["Title"].Value<string>() == Constants.SITEFOOTER_TITLENODEKEY);

            // Ensure the title node contains the expected child elements
            if(titleNode == null || titleNode["Nodes"] == null || titleNode["Nodes"].Count() == 0 || titleNode["Nodes"][0]["Title"] == null)
            {
                // The expected child elements were not found
                return null;
            }

            // Retrieve the title
            var title = titleNode["Nodes"][0]["Title"].Value<string>();
            return title;
        }

        /// <summary>
        /// Sets the title shown in the footer
        /// </summary>
        /// <param name="web">Web instance to set the footer title of</param>
        /// <param name="title">Title to show in the footer</param>
        /// <returns>Boolean indicating if setting the title succeeded</returns>
        public static bool SetFooterTitle(this Web web, string title)
        {
            web.EnsureProperty(w => w.ServerRelativeUrl);
            var responseString = web.ExecutePostAsync("/_api/navigation/SaveMenuState", 
                                                @"{""menuState"":{""StartingNodeTitle"":""" + Constants.SITEFOOTER_NODEKEY + @""",""SPSitePrefix"":""/"",""SPWebPrefix"":""" + web.ServerRelativeUrl + @""",""FriendlyUrlPrefix"":"""",""SimpleUrl"":"""",""Nodes"":[{""NodeType"":0,""Title"":""" + Constants.SITEFOOTER_TITLENODEKEY + @""",""Key"":""2004"",""FriendlyUrlSegment"":"""",""Nodes"":[{""NodeType"":0,""Title"":""" + title + @""",""FriendlyUrlSegment"":""""}]}]}}").GetAwaiter().GetResult();
            var responseJson = JObject.Parse(responseString);
            var requestSucceeded = responseJson != null && responseJson["value"] != null && responseJson["value"].Value<string>() == "200";
            return requestSucceeded;
        }

        /// <summary>
        /// Returns the server relative URL of the logo shown in the footer
        /// </summary>
        /// <param name="web">Web instance to return the footer logo url of</param>
        /// <returns>Server relative URL of the logo shown in the footer or NULL if no footer has been set</returns>
        public static string GetFooterLogoUrl(this Web web)
        {
            var structureString = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'").GetAwaiter().GetResult();
            var menuState = JObject.Parse(structureString);

            if (menuState["Nodes"] == null)
            {
                // No information is returned which helps us to identity the logo node
                return null;
            }

            // Retrieve the Node representing the logo node
            var logoUrlNode = menuState["Nodes"].FirstOrDefault(n => n["Title"].Value<string>() == Constants.SITEFOOTER_LOGONODEKEY);

            // Ensure the logo node contains the expected child elements
            if (logoUrlNode == null || logoUrlNode["SimpleUrl"] == null)
            {
                // The expected child elements were not found
                return null;
            }

            // Retrieve the logo url
            var logoUrl = logoUrlNode["SimpleUrl"].Value<string>();
            return logoUrl;
        }

        /// <summary>
        /// Sets the logo shown in the footer
        /// </summary>
        /// <param name="web">Web instance to set the footer logo url of</param>
        /// <param name="logoUrl">Server relative path to the logo to show in the footer</param>
        /// <returns>Boolean indicating if setting the logo succeeded</returns>
        public static bool SetFooterLogoUrl(this Web web, string logoUrl)
        {
            web.EnsureProperty(w => w.ServerRelativeUrl);
            var responseString = web.ExecutePostAsync("/_api/navigation/SaveMenuState",
                                                @"{""menuState"":{""StartingNodeTitle"":""" + Constants.SITEFOOTER_NODEKEY + @""",""SPSitePrefix"":""/"",""SPWebPrefix"":""" + web.ServerRelativeUrl + @""",""FriendlyUrlPrefix"":"""",""SimpleUrl"":"""",""Nodes"":[{""NodeType"":0,""Title"":""" + Constants.SITEFOOTER_LOGONODEKEY + @""",""Key"":""2006"",""SimpleUrl"":""" + logoUrl + @""",""FriendlyUrlSegment"":""""}]}}").GetAwaiter().GetResult();
            var responseJson = JObject.Parse(responseString);
            var requestSucceeded = responseJson != null && responseJson["value"] != null && responseJson["value"].Value<string>() == "200";
            return requestSucceeded;
        }

        /// <summary>
        /// Removes the logo shown in the footer
        /// </summary>
        /// <param name="web">Web instance to remove the footer of</param>
        /// <returns>Boolean indicating if removing the logo succeeded</returns>
        public static bool RemoveFooterLogoUrl(this Web web)
        {
            web.EnsureProperty(w => w.ServerRelativeUrl);
            var responseString = web.ExecutePostAsync("/_api/navigation/SaveMenuState",
                                                @"{""menuState"":{""StartingNodeTitle"":""" + Constants.SITEFOOTER_NODEKEY + @""",""SPSitePrefix"":""/"",""SPWebPrefix"":""" + web.ServerRelativeUrl + @""",""FriendlyUrlPrefix"":"""",""SimpleUrl"":"""",""Nodes"":[{""NodeType"":0,""Title"":""" + Constants.SITEFOOTER_LOGONODEKEY + @""",""IsDeleted"":""True"",""FriendlyUrlSegment"":""""}]}}").GetAwaiter().GetResult();
            var responseJson = JObject.Parse(responseString);
            var requestSucceeded = responseJson != null && responseJson["value"] != null && responseJson["value"].Value<string>() == "200";
            return requestSucceeded;
        }
#endif
        #endregion

        #region Custom actions
        /// <summary>
        /// Adds custom action to a web. If the CustomAction exists the item will be updated.
        /// Setting CustomActionEntity.Remove == true will delete the CustomAction.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="customAction">Information about the custom action be added or deleted</param>
        /// <example>
        /// var editAction = new CustomActionEntity()
        /// {
        ///    Title = "Edit Site Classification",
        ///    Description = "Manage business impact information for site collection or sub sites.",
        ///    Sequence = 1000,
        ///    Group = "SiteActions",
        ///    Location = "Microsoft.SharePoint.StandardMenu",
        ///    Url = EditFormUrl,
        ///    ImageUrl = EditFormImageUrl,
        ///    Rights = new BasePermissions(),
        /// };
        /// editAction.Rights.Set(PermissionKind.ManageWeb);
        /// web.AddCustomAction(editAction);
        /// </example>
        /// <returns>True if action was successfull</returns>
        public static bool AddCustomAction(this Web web, CustomActionEntity customAction)
        {
            return AddCustomActionImplementation(web, customAction);
        }

        /// <summary>
        /// Adds custom action to a site collection. If the CustomAction exists the item will be updated.
        /// Setting CustomActionEntity.Remove == true will delete the CustomAction.
        /// </summary>
        /// <param name="site">Site collection to be processed</param>
        /// <param name="customAction">Information about the custom action be added or deleted</param>
        /// <returns>True if action was successfull</returns>
        public static bool AddCustomAction(this Site site, CustomActionEntity customAction)
        {
            return AddCustomActionImplementation(site, customAction);
        }

        private static bool AddCustomActionImplementation(ClientObject clientObject, CustomActionEntity customAction)
        {
            UserCustomAction targetAction;
            UserCustomActionCollection existingActions;
            if (clientObject is Web)
            {
                var web = (Web)clientObject;

                existingActions = web.UserCustomActions;
                web.Context.Load(existingActions);
                web.Context.ExecuteQueryRetry();

                targetAction = web.UserCustomActions.FirstOrDefault(uca => uca.Name == customAction.Name);
            }
            else
            {
                var site = (Site)clientObject;

                existingActions = site.UserCustomActions;
                site.Context.Load(existingActions);
                site.Context.ExecuteQueryRetry();

                targetAction = site.UserCustomActions.FirstOrDefault(uca => uca.Name == customAction.Name);
            }

            if (targetAction == null)
            {
                // If we're removing the custom action then we need to leave when not found...else we're creating the custom action
                if (customAction.Remove)
                {
                    return true;
                }
                targetAction = existingActions.Add();
            }
            else if (customAction.Remove)
            {
                targetAction.DeleteObject();
                clientObject.Context.ExecuteQueryRetry();
                return true;
            }

            targetAction.Name = customAction.Name;
            targetAction.Description = customAction.Description;
            targetAction.Location = customAction.Location;
            targetAction.Sequence = customAction.Sequence;
#if !SP2013 && !SP2016
            targetAction.ClientSideComponentId = customAction.ClientSideComponentId;
            targetAction.ClientSideComponentProperties = customAction.ClientSideComponentProperties;
#endif
#if !ONPREMISES
            targetAction.HostProperties = customAction.ClientSideHostProperties;
#endif
            if (customAction.Location == JavaScriptExtensions.SCRIPT_LOCATION)
            {
                targetAction.ScriptBlock = customAction.ScriptBlock;
                targetAction.ScriptSrc = customAction.ScriptSrc;
            }
            else
            {
                targetAction.Url = customAction.Url;
                targetAction.Group = customAction.Group;
                targetAction.Title = customAction.Title;
                targetAction.ImageUrl = customAction.ImageUrl;

                if (customAction.RegistrationId != null)
                {
                    targetAction.RegistrationId = customAction.RegistrationId;
                }

                if (!string.IsNullOrEmpty(customAction.CommandUIExtension))
                {
                    targetAction.CommandUIExtension = customAction.CommandUIExtension;
                }

                if (customAction.Rights != null)
                {
                    targetAction.Rights = customAction.Rights;
                }

                if (customAction.RegistrationType.HasValue)
                {
                    targetAction.RegistrationType = customAction.RegistrationType.Value;
                }
            }

            targetAction.Update();
            if (clientObject is Web)
            {
                var web = (Web)clientObject;
                web.Context.Load(web, w => w.UserCustomActions);
                web.Context.ExecuteQueryRetry();
            }
            else
            {
                var site = (Site)clientObject;
                site.Context.Load(site, s => s.UserCustomActions);
                site.Context.ExecuteQueryRetry();
            }

            return true;
        }


        /// <summary>
        /// Returns all custom actions in a web
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>Returns all custom actions</returns>
        public static IEnumerable<UserCustomAction> GetCustomActions(this Web web, params Expression<Func<UserCustomAction, object>>[] expressions)
        {
            var clientContext = (ClientContext)web.Context;

            List<UserCustomAction> actions = new List<UserCustomAction>();

            if (expressions != null && expressions.Any())
            {
                clientContext.Load(web.UserCustomActions, u => u.IncludeWithDefaultProperties(expressions));
            }
            else
            {
                clientContext.Load(web.UserCustomActions);
            }
            clientContext.ExecuteQueryRetry();

            foreach (UserCustomAction uca in web.UserCustomActions)
            {
                actions.Add(uca);
            }
            return actions;
        }

        /// <summary>
        /// Returns all custom actions in a web
        /// </summary>
        /// <param name="site">The site to process</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>Returns all custom actions</returns>
        public static IEnumerable<UserCustomAction> GetCustomActions(this Site site, params Expression<Func<UserCustomAction, object>>[] expressions)
        {
            var clientContext = (ClientContext)site.Context;

            List<UserCustomAction> actions = new List<UserCustomAction>();
            if (expressions != null && expressions.Any())
            {
                clientContext.Load(site.UserCustomActions, u => u.IncludeWithDefaultProperties(expressions));
            }
            else
            {
                clientContext.Load(site.UserCustomActions);
            }
            clientContext.ExecuteQueryRetry();

            foreach (UserCustomAction uca in site.UserCustomActions)
            {
                actions.Add(uca);
            }
            return actions;
        }

        /// <summary>
        /// Removes a custom action
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="id">The id of the action to remove. <seealso>
        ///         <cref>GetCustomActions</cref>
        ///     </seealso>
        /// </param>
        public static void DeleteCustomAction(this Web web, Guid id)
        {
            var clientContext = (ClientContext)web.Context;

            clientContext.Load(web.UserCustomActions);
            clientContext.ExecuteQueryRetry();

            foreach (UserCustomAction action in web.UserCustomActions)
            {
                if (action.Id == id)
                {
                    action.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                    break;
                }
            }
        }

        /// <summary>
        /// Removes a custom action
        /// </summary>
        /// <param name="site">The site to process</param>
        /// <param name="id">The id of the action to remove. <seealso>
        ///         <cref>GetCustomActions</cref>
        ///     </seealso>
        /// </param>
        public static void DeleteCustomAction(this Site site, Guid id)
        {
            var clientContext = (ClientContext)site.Context;

            clientContext.Load(site.UserCustomActions);
            clientContext.ExecuteQueryRetry();

            foreach (UserCustomAction action in site.UserCustomActions)
            {
                if (action.Id == id)
                {
                    action.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                    break;
                }
            }
        }

        /// <summary>
        /// Utility method to check particular custom action already exists on the web
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="name">Name of the custom action</param>
        /// <returns></returns>        
        public static bool CustomActionExists(this Web web, string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name));

            web.Context.Load(web.UserCustomActions);
            web.Context.ExecuteQueryRetry();

            var customActions = web.UserCustomActions.AsEnumerable();
            foreach (var customAction in customActions)
            {
                var customActionName = customAction.Name;
                if (!string.IsNullOrEmpty(customActionName) &&
                    customActionName.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Utility method to check particular custom action already exists on the web
        /// </summary>
        /// <param name="site">Site to process</param>
        /// <param name="name">Name of the custom action</param>
        /// <returns></returns>        
        public static bool CustomActionExists(this Site site, string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name));

            site.Context.Load(site.UserCustomActions);
            site.Context.ExecuteQueryRetry();

            var customActions = site.UserCustomActions.AsEnumerable();
            foreach (var customAction in customActions)
            {
                var customActionName = customAction.Name;
                if (!string.IsNullOrEmpty(customActionName) &&
                    customActionName.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
    }

    /// <summary>
    /// Defines the kind of Managed Navigation for a site
    /// </summary>
    public enum ManagedNavigationKind
    {
        /// <summary>
        /// Current Navigation
        /// </summary>
        Current,
        /// <summary>
        /// Global Navigation
        /// </summary>
        Global
    }
}
