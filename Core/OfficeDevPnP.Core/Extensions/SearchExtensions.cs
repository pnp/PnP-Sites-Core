using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client.Search.Administration;
using Microsoft.SharePoint.Client.Search.Portability;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Text;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for Search extension methods
    /// </summary>
    public static partial class SearchExtensions
    {
        /// <summary>
        /// Exports the search settings to file.
        /// </summary>
        /// <param name="context">Context for SharePoint objects and operations</param>
        /// <param name="exportFilePath">Path where to export the search settings</param>
        /// <param name="searchSettingsExportLevel">Search settings export level
        /// Reference: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.administration.searchobjectlevel(v=office.15).aspx
        /// </param>
        public static void ExportSearchSettings(this ClientContext context, string exportFilePath, SearchObjectLevel searchSettingsExportLevel)
        {
            if (string.IsNullOrEmpty(exportFilePath))
            {
                throw new ArgumentNullException(nameof(exportFilePath));
            }

            var searchConfig = GetSearchConfigurationImplementation(context, searchSettingsExportLevel);

            if (searchConfig != null)
            {
                System.IO.File.WriteAllText(exportFilePath, searchConfig, Encoding.ASCII);
            }
            else
            {
                throw new Exception("No search settings to export.");
            }
        }

        /// <summary>
        /// Returns the current search configuration as as string
        /// </summary>
        /// <param name="web">A SharePoint site/subsiste</param>
        /// <returns>Returns search configuration</returns>
        public static string GetSearchConfiguration(this Web web)
        {
            return GetSearchConfigurationImplementation(web.Context, SearchObjectLevel.SPWeb);
        }

        /// <summary>
        /// Returns the current search configuration as as string
        /// </summary>
        /// <param name="site">A SharePoint site</param>
        /// <returns>Returns search configuration</returns>
        public static string GetSearchConfiguration(this Site site)
        {
            return GetSearchConfigurationImplementation(site.Context, SearchObjectLevel.SPSite);
        }

        /// <summary>
        /// Returns the current search configuration for the specified object level
        /// </summary>
        /// <param name="context">ClinetRuntimeContext for SharePoint objects and operations</param>
        /// <param name="searchSettingsObjectLevel">A site server level value. i.e, SPWeb/SPSite/SPSiteSubscription/Ssa</param>
        /// <returns>Returns search configuration</returns>
        private static string GetSearchConfigurationImplementation(ClientRuntimeContext context, SearchObjectLevel searchSettingsObjectLevel)
        {
            SearchConfigurationPortability sconfig = new SearchConfigurationPortability(context);
            SearchObjectOwner owner = new SearchObjectOwner(context, searchSettingsObjectLevel);

            ClientResult<string> configresults = sconfig.ExportSearchConfiguration(owner);
            context.ExecuteQueryRetry();

            return configresults.Value;
        }

        /// <summary>
        /// Imports search settings from file.
        /// </summary>
        /// <param name="context">Context for SharePoint objects and operations</param>
        /// <param name="searchSchemaImportFilePath">Search schema xml file path</param>
        /// <param name="searchSettingsImportLevel">Search settings import level
        /// Reference: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.administration.searchobjectlevel(v=office.15).aspx
        /// </param>
        public static void ImportSearchSettings(this ClientContext context, string searchSchemaImportFilePath, SearchObjectLevel searchSettingsImportLevel)
        {
            if (string.IsNullOrEmpty(searchSchemaImportFilePath))
            {
                throw new ArgumentNullException(nameof(searchSchemaImportFilePath));
            }
            SetSearchConfigurationImplementation(context, searchSettingsImportLevel, System.IO.File.ReadAllText(searchSchemaImportFilePath));
        }

        /// <summary>
        /// Imports search settings from configuration xml.
        /// </summary>
        /// <param name="context">Context for SharePoint objects and operations</param>
        /// <param name="searchConfiguration">Search schema xml file path</param>
        /// <param name="searchSettingsImportLevel">Search settings import level
        /// Reference: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.administration.searchobjectlevel(v=office.15).aspx
        /// </param>
        public static void ImportSearchSettingsConfiguration(this ClientContext context, string searchConfiguration, SearchObjectLevel searchSettingsImportLevel)
        {
            if (string.IsNullOrEmpty(searchConfiguration))
            {
                throw new ArgumentNullException(nameof(searchConfiguration));
            }
            SetSearchConfigurationImplementation(context, searchSettingsImportLevel, searchConfiguration);
        }


        /// <summary>
        /// Delete search settings from configuration xml.
        /// </summary>
        /// <param name="context">Context for SharePoint objects and operations</param>
        /// <param name="searchConfiguration">Search schema xml file path</param>
        /// <param name="searchSettingsImportLevel">Search settings import level
        /// Reference: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.administration.searchobjectlevel(v=office.15).aspx
        /// </param>
        public static void DeleteSearchSettings(this ClientContext context, string searchConfiguration, SearchObjectLevel searchSettingsImportLevel)
        {
            if (string.IsNullOrEmpty(searchConfiguration))
            {
                throw new ArgumentNullException(nameof(searchConfiguration));
            }

            DeleteSearchConfigurationImplementation(context, searchSettingsImportLevel, searchConfiguration);

        }

        /// <summary>
        /// Sets the search configuration
        /// </summary>
        /// <param name="web">A SharePoint site/subsite</param>
        /// <param name="searchConfiguration">search configuration</param>
        public static void SetSearchConfiguration(this Web web, string searchConfiguration)
        {
            SetSearchConfigurationImplementation(web.Context, SearchObjectLevel.SPWeb, searchConfiguration);
        }

        /// <summary>
        /// Sets the search configuration
        /// </summary>
        /// <param name="site">A SharePoint site</param>
        /// <param name="searchConfiguration">search configuration</param>
        public static void SetSearchConfiguration(this Site site, string searchConfiguration)
        {
            SetSearchConfigurationImplementation(site.Context, SearchObjectLevel.SPSite, searchConfiguration);
        }

        /// <summary>
        /// Sets the search configuration at the specified object level
        /// </summary>
        /// <param name="context"></param>
        /// <param name="searchObjectLevel"></param>
        /// <param name="searchConfiguration"></param>
        private static void SetSearchConfigurationImplementation(ClientRuntimeContext context, SearchObjectLevel searchObjectLevel, string searchConfiguration)
        {
#if ONPREMISES
            if (searchObjectLevel == SearchObjectLevel.Ssa)
            {
                // Reference: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.portability.searchconfigurationportability_members.aspx
                throw new Exception("You cannot import customized search configuration settings to a Search service application (SSA).");
            }
#endif
            SearchConfigurationPortability searchConfig = new SearchConfigurationPortability(context);
            SearchObjectOwner owner = new SearchObjectOwner(context, searchObjectLevel);

            // Import search configuration
            searchConfig.ImportSearchConfiguration(owner, searchConfiguration);
            context.Load(searchConfig);
            context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Delete the search configuration - does not apply to managed properties.
        /// </summary>
        /// <param name="web">A SharePoint site/subsite</param>
        /// <param name="searchConfiguration">search configuration</param>
        public static void DeleteSearchConfiguration(this Web web, string searchConfiguration)
        {
            DeleteSearchConfigurationImplementation(web.Context, SearchObjectLevel.SPWeb, searchConfiguration);
        }

        /// <summary>
        /// Delete the search configuration - does not apply to managed properties.
        /// </summary>
        /// <param name="site">A SharePoint site</param>
        /// <param name="searchConfiguration">search configuration</param>
        public static void DeleteSearchConfiguration(this Site site, string searchConfiguration)
        {
            DeleteSearchConfigurationImplementation(site.Context, SearchObjectLevel.SPSite, searchConfiguration);
        }

        /// <summary>
        /// Delete the search configuration at the specified object level - does not apply to managed properties.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="searchObjectLevel"></param>
        /// <param name="searchConfiguration"></param>
        private static void DeleteSearchConfigurationImplementation(ClientRuntimeContext context, SearchObjectLevel searchObjectLevel, string searchConfiguration)
        {
#if ONPREMISES
            if (searchObjectLevel == SearchObjectLevel.Ssa)
            {
                // Reference: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.portability.searchconfigurationportability_members.aspx
                throw new Exception("You cannot import customized search configuration settings to a Search service application (SSA).");
            }
#endif
            SearchConfigurationPortability searchConfig = new SearchConfigurationPortability(context);
            SearchObjectOwner owner = new SearchObjectOwner(context, searchObjectLevel);

            // Delete search configuration
            searchConfig.DeleteSearchConfiguration(owner, searchConfiguration);
            context.Load(searchConfig);
            context.ExecuteQueryRetry();
        }

#if !ONPREMISES
        public static void SetSearchBoxPlaceholderText(this Site site, string placeholderText)
        {
            SetSearchBoxPlaceholderTextImpl(site.RootWeb, placeholderText, true);
        }

        public static void SetSearchBoxPlaceholderText(this Web web, string placeholderText)
        {
            SetSearchBoxPlaceholderTextImpl(web, placeholderText, false);
        }

        private static void SetSearchBoxPlaceholderTextImpl(Web web, string placeholderText, bool siteScope)
        {
            if (placeholderText == null)
            {
                throw new ArgumentNullException(nameof(placeholderText));
            }

            ClientContext adminContext = null;

            var ctx = ((ClientContext)web.Context);
            ctx.Load(ctx.Web, w => w.EffectiveBasePermissions); // reload permissions
            ctx.ExecuteQueryRetry();

            Site site = ctx.Site;

            #region Enable scripting if needed and context has access
            Tenant tenant = null;
            if (ctx.Web.IsNoScriptSite() && TenantExtensions.IsCurrentUserTenantAdmin(ctx))
            {
                site.EnsureProperty(s => s.Url);
                var adminSiteUrl = web.GetTenantAdministrationUrl();
                adminContext = ctx.Clone(adminSiteUrl);
                tenant = new Tenant(adminContext);
                tenant.SetSiteProperties(site.Url, noScriptSite: false);
            }
            #endregion

            try
            {
                if (siteScope)
                {
                    site.SearchBoxPlaceholderText = placeholderText;
                }
                else
                {
                    web.SearchBoxPlaceholderText = placeholderText;
                    web.Update();
                }
                ctx.ExecuteQueryRetry();
            }
            finally
            {
                #region Disable scripting if previously enabled
                if (adminContext != null)
                {
                    // Reset disabling setting the property bag if needed
                    tenant.SetSiteProperties(site.Url, noScriptSite: true);
                    adminContext.Dispose();
                }
                #endregion
            }
        }
#endif

        /// <summary>
        /// Sets the search center URL on site collection (Site Settings -> Site collection administration --> Search Settings)
        /// </summary>
        /// <param name="web">SharePoint site - root web</param>
        /// <param name="searchCenterUrl">Search center URL</param>
        public static void SetSiteCollectionSearchCenterUrl(this Web web, string searchCenterUrl)
        {
            if (searchCenterUrl == null)
            {
                throw new ArgumentNullException(nameof(searchCenterUrl));
            }

            // Currently there is no direct API available to set the search center URL on web.
            // Set search setting at site level   

#if !ONPREMISES
            #region Enable scripting if needed and context has access
            Tenant tenant = null;
            Site site = null;
            ClientContext adminContext = null;
            var ctx = ((ClientContext)web.Context);
            ctx.Load(ctx.Web, w => w.EffectiveBasePermissions); // reload permissions in case changed after connect
            ctx.ExecuteQueryRetry();

            if (ctx.Web.IsNoScriptSite() && TenantExtensions.IsCurrentUserTenantAdmin(web.Context as ClientContext))
            {
                site = ((ClientContext)web.Context).Site;
                site.EnsureProperty(s => s.Url);

                var adminSiteUrl = web.GetTenantAdministrationUrl();
                adminContext = web.Context.Clone(adminSiteUrl);
                tenant = new Tenant(adminContext);
                tenant.SetSiteProperties(site.Url, noScriptSite: false);
            }
            #endregion
#endif

            try
            {
                // if another value was set then respect that
                if (String.IsNullOrEmpty(web.GetPropertyBagValueString("SRCH_SB_SET_SITE", string.Empty)))
                {
                    web.SetPropertyBagValue("SRCH_SB_SET_SITE", "{'Inherit':false,'ResultsPageAddress':null,'ShowNavigation':true}");
                }

                if (!string.IsNullOrEmpty(searchCenterUrl))
                {
                    // Set search center URL
                    web.SetPropertyBagValue("SRCH_ENH_FTR_URL_SITE", searchCenterUrl);
                }
                else
                {
                    // When search center URL is blank remove the property (like the SharePoint UI does)
                    web.RemovePropertyBagValue("SRCH_ENH_FTR_URL_SITE");
                }
            }
            finally
            {
#if !ONPREMISES
                #region Disable scripting if previously enabled
                if (adminContext != null)
                {
                    // Reset disabling setting the property bag if needed
                    tenant.SetSiteProperties(site.Url, noScriptSite: true);
                    adminContext.Dispose();
                }
                #endregion
#endif
            }
        }

        /// <summary>
        /// Get the search center URL for the site collection (Site Settings -> Site collection administration --> Search Settings)
        /// </summary>
        /// <param name="web">SharePoint site - root web</param>
        /// <returns>Search center URL for web</returns>
        public static string GetSiteCollectionSearchCenterUrl(this Web web)
        {
            // Currently there is no direct API available to get the search center URL on web.
            // Get search center URL
            return web.GetPropertyBagValueString("SRCH_ENH_FTR_URL_SITE", string.Empty);
        }

        /// <summary>
        /// Sets the search results page URL on current web (Site Settings -> Search --> Search Settings)
        /// </summary>
        /// <param name="web">SharePoint current web</param>
        /// <param name="searchCenterUrl">Search results page URL</param>
        public static void SetWebSearchCenterUrl(this Web web, string searchCenterUrl)
        {
            if (searchCenterUrl == null)
            {
                throw new ArgumentNullException(nameof(searchCenterUrl));
            }

            // Currently there is no direct API available to set the search center URL on web.
            // Set search setting at web level   

#if !ONPREMISES
            #region Enable scripting if needed and context has access
            Tenant tenant = null;
            Site site = null;
            ClientContext adminContext = null;
            if (web.IsNoScriptSite() && TenantExtensions.IsCurrentUserTenantAdmin(web.Context as ClientContext))
            {
                site = ((ClientContext)web.Context).Site;
                site.EnsureProperty(s => s.Url);

                var adminSiteUrl = web.GetTenantAdministrationUrl();
                adminContext = web.Context.Clone(adminSiteUrl);
                tenant = new Tenant(adminContext);
                tenant.SetSiteProperties(site.Url, noScriptSite: false);
            }
            #endregion
#endif

            try
            {
                string keyName = web.IsSubSite() ? "SRCH_SB_SET_WEB" : "SRCH_SB_SET_SITE";

                if (!string.IsNullOrEmpty(searchCenterUrl))
                {
                    // Set search results page URL
                    web.SetPropertyBagValue(keyName, "{\"Inherit\":false,\"ResultsPageAddress\":\"" + searchCenterUrl + "\",\"ShowNavigation\":false}");
                }
                else
                {
                    // When search results page URL is blank remove the property (like the SharePoint UI does)
                    web.RemovePropertyBagValue(keyName);
                }
            }
            catch (ServerUnauthorizedAccessException e)
            {
                const string errorMsg = "For modern sites you need to be a SharePoint admin when setting the search redirect URL programatically.\n\nPlease use the classic UI at '/_layouts/15/enhancedSearch.aspx?level=sitecol'.";
                Log.Error(e, Constants.LOGGING_SOURCE, errorMsg);
                throw new ApplicationException(errorMsg, e);
            }
            finally
            {
#if !ONPREMISES

                #region Disable scripting if previously enabled
                if (adminContext != null)
                {
                    // Reset disabling setting the property bag if needed
                    tenant.SetSiteProperties(site.Url, noScriptSite: true);
                    adminContext.Dispose();
                }
                #endregion
#endif
            }
        }

        /// <summary>
        /// Get the search results page URL for the web (Site Settings -> Search --> Search Settings)
        /// </summary>
        /// <param name="web">SharePoint site - current web</param>
        /// <param name="urlOnly">Allows to declare to return the URL only and not the full JSON settings</param>
        /// <returns>Search results page URL for web</returns>
        public static string GetWebSearchCenterUrl(this Web web, bool urlOnly = false)
        {
            bool isSubSite = web.IsSubSite();
            string keyName = isSubSite ? "SRCH_SB_SET_WEB" : "SRCH_SB_SET_SITE";

            // Get the Search Settings JSON value
            var searchSettingsValue = web.GetPropertyBagValueString(keyName, string.Empty);
            if (!isSubSite && string.IsNullOrWhiteSpace(searchSettingsValue))
            {
                // fallback to read web value on sc root
                searchSettingsValue = web.GetPropertyBagValueString("SRCH_SB_SET_WEB", string.Empty);
            }

            // Convert the settings into a typed object
            var searchSettings = JsonConvert.DeserializeAnonymousType(searchSettingsValue, new
            {
                Inherit = false,
                ResultsPageAddress = String.Empty,
                ShowNavigation = false,
            });

            if (searchSettings != null && !searchSettings.Inherit)
            {
                if (!urlOnly)
                {
                    // Return the whole JSON settings
                    return searchSettingsValue;
                }
                else
                {
                    // Return the search results page URL of the current web
                    return searchSettings?.ResultsPageAddress;
                }
            }
            else
            {
                // If we're inheriting settings, just return NULL
                return null;
            }
        }
    }
}
