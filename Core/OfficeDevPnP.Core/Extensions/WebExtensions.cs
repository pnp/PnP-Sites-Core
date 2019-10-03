using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.Search.Query;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using System.Reflection;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that deals with site (both site collection and web site) creation, status, retrieval and settings
    /// </summary>
    public static partial class WebExtensions
    {
        const string MSG_CONTEXT_CLOSED = "ClientContext gets closed after action is completed. Calling ExecuteQuery again returns an error. Verify that you have an open ClientContext object.";
        const string SITE_STATUS_ACTIVE = "Active";
        const string SITE_STATUS_CREATING = "Creating";
        const string SITE_STATUS_RECYCLED = "Recycled";
        const string INDEXED_PROPERTY_KEY = "vti_indexedpropertykeys";

        #region Web (site) query, creation and deletion

        /// <summary>
        /// Returns the Base Template ID for the current web
        /// </summary>
        /// <param name="parentWeb">The parent Web (site) to get the base template from</param>
        /// <returns>The Base Template ID for the current web</returns>
        public static String GetBaseTemplateId(this Web parentWeb)
        {
            parentWeb.EnsureProperties(w => w.WebTemplate, w => w.Configuration);
            return ($"{parentWeb.WebTemplate}#{parentWeb.Configuration}");
        }

        /// <summary>
        /// Adds a new child Web (site) to a parent Web.
        /// </summary>
        /// <param name="parentWeb">The parent Web (site) to create under</param>
        /// <param name="subsite">Details of the Web (site) to add. Only Title, Url (as the leaf URL), Description, Template and Language are used.</param>
        /// <param name="inheritPermissions">Specifies whether the new site will inherit permissions from its parent site.</param>
        /// <param name="inheritNavigation">Specifies whether the site inherits navigation.</param>
        /// <returns></returns>
        public static Web CreateWeb(this Web parentWeb, SiteEntity subsite, bool inheritPermissions = true, bool inheritNavigation = true)
        {
            return CreateWeb(parentWeb, subsite.Title, subsite.Url, subsite.Description, subsite.Template, (int)subsite.Lcid, inheritPermissions, inheritNavigation);
        }

        /// <summary>
        /// Adds a new child Web (site) to a parent Web.
        /// </summary>
        /// <param name="parentWeb">The parent Web (site) to create under</param>
        /// <param name="title">The title of the new site. </param>
        /// <param name="leafUrl">A string that represents the URL leaf name.</param>
        /// <param name="description">The description of the new site. </param>
        /// <param name="template">The name of the site template to be used for creating the new site. </param>
        /// <param name="language">The locale ID that specifies the language of the new site. </param>
        /// <param name="inheritPermissions">Specifies whether the new site will inherit permissions from its parent site.</param>
        /// <param name="inheritNavigation">Specifies whether the site inherits navigation.</param>
        public static Web CreateWeb(this Web parentWeb, string title, string leafUrl, string description, string template, int language, bool inheritPermissions = true, bool inheritNavigation = true)
        {
            if (leafUrl.ContainsInvalidUrlChars())
            {
                throw new ArgumentException("The argument must be a single web URL and cannot contain path characters.", nameof(leafUrl));
            }

            bool isNoScript = parentWeb.IsNoScriptSite();

            Log.Info(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_CreateWeb, leafUrl, template);
            WebCreationInformation creationInfo = new WebCreationInformation
            {
                Url = leafUrl,
                Title = title,
                Description = description,
                UseSamePermissionsAsParentSite = inheritPermissions,
                WebTemplate = template,
                Language = language
            };

            Web newWeb = parentWeb.Webs.Add(creationInfo);
            parentWeb.Context.ExecuteQueryRetry();

            if (!isNoScript)
            {
                newWeb.Navigation.UseShared = inheritNavigation;
            }
            newWeb.Update();

            parentWeb.Context.ExecuteQueryRetry();

            return newWeb;
        }

        /// <summary>
        /// Deletes the child website with the specified leaf URL, from a parent Web, if it exists.
        /// </summary>
        /// <param name="parentWeb">The parent Web (site) to delete from</param>
        /// <param name="leafUrl">A string that represents the URL leaf name.</param>
        /// <returns>true if the web was deleted; otherwise false if nothing was done</returns>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static bool DeleteWeb(this Web parentWeb, string leafUrl)
        {
            if (leafUrl.ContainsInvalidUrlChars())
            {
                throw new ArgumentException("The argument must be a single web URL and cannot contain path characters.", nameof(leafUrl));
            }

            var deleted = false;
            parentWeb.EnsureProperties(w => w.ServerRelativeUrl);

            var serverRelativeUrl = UrlUtility.Combine(parentWeb.ServerRelativeUrl, leafUrl);
            var webs = parentWeb.Webs;
            // NOTE: Predicate does not take into account a required case-insensitive comparison
            //var results = parentWeb.Context.LoadQuery<Web>(webs.Where(item => item.ServerRelativeUrl == serverRelativeUrl));
            parentWeb.Context.Load(webs, wc => wc.Include(w => w.ServerRelativeUrl));
            parentWeb.Context.ExecuteQueryRetry();
            var existingWeb = webs.FirstOrDefault(item => string.Equals(item.ServerRelativeUrl, serverRelativeUrl, StringComparison.OrdinalIgnoreCase));
            if (existingWeb != null)
            {
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_DeleteWeb, serverRelativeUrl);
                existingWeb.DeleteObject();
                parentWeb.Context.ExecuteQueryRetry();
                deleted = true;
            }
            else
            {
                Log.Debug(Constants.LOGGING_SOURCE, "Delete requested but web '{0}' not found, nothing to do.", serverRelativeUrl);
            }
            return deleted;
        }

        /// <summary>
        /// Gets the collection of the URLs of all Web sites that are contained within the site collection,
        /// including the top-level site and its subsites.
        /// </summary>
        /// <param name="site">Site collection to retrieve the URLs for.</param>
        /// <returns>An enumeration containing the full URLs as strings.</returns>
        /// <remarks>
        /// <para>
        /// This is analagous to the <code>SPSite.AllWebs</code> property and can be used to get a collection
        /// of all web site URLs to loop through, e.g. for branding.
        /// </para>
        /// </remarks>
        public static IEnumerable<string> GetAllWebUrls(this Site site)
        {
            var siteContext = site.Context;
            siteContext.Load(site, s => s.Url);
            siteContext.ExecuteQueryRetry();
            var queue = new Queue<string>();
            queue.Enqueue(site.Url);
            while (queue.Count > 0)
            {
                var currentUrl = queue.Dequeue();
                using (var webContext = siteContext.Clone(currentUrl))
                {
                    webContext.Load(webContext.Web, web => web.Webs);
                    webContext.ExecuteQueryRetry();
                    foreach (var subWeb in webContext.Web.Webs)
                    {
                        queue.Enqueue(subWeb.Url);
                    }
                }
                yield return currentUrl;
            }
        }

        /// <summary>
        /// Returns the child Web site with the specified leaf URL.
        /// </summary>
        /// <param name="parentWeb">The Web site to check under</param>
        /// <param name="leafUrl">A string that represents the URL leaf name.</param>
        /// <returns>The requested Web, if it exists, otherwise null.</returns>
        /// <remarks>
        /// <para>
        /// The ServerRelativeUrl property of the retrieved Web is instantiated.
        /// </para>
        /// </remarks>
        public static Web GetWeb(this Web parentWeb, string leafUrl)
        {
            if (leafUrl.ContainsInvalidUrlChars())
            {
                throw new ArgumentException("The argument must be a single web URL and cannot contain path characters.", nameof(leafUrl));
            }

            parentWeb.EnsureProperty(w => w.ServerRelativeUrl);

            var serverRelativeUrl = UrlUtility.Combine(parentWeb.ServerRelativeUrl, leafUrl);
            var webs = parentWeb.Webs;
            // NOTE: Predicate does not take into account a required case-insensitive comparison
            //var results = parentWeb.Context.LoadQuery<Web>(webs.Where(item => item.ServerRelativeUrl == serverRelativeUrl));
            parentWeb.Context.Load(webs, wc => wc.Include(w => w.ServerRelativeUrl));
            parentWeb.Context.ExecuteQueryRetry();
            var childWeb = webs.FirstOrDefault(item => string.Equals(item.ServerRelativeUrl, serverRelativeUrl, StringComparison.OrdinalIgnoreCase));
            return childWeb;
        }

        /// <summary>
        /// Determines if a child Web site with the specified leaf URL exists.
        /// </summary>
        /// <param name="parentWeb">The Web site to check under</param>
        /// <param name="leafUrl">A string that represents the URL leaf name.</param>
        /// <returns>true if the Web (site) exists; otherwise false</returns>
        public static bool WebExists(this Web parentWeb, string leafUrl)
        {
            if (leafUrl.ContainsInvalidUrlChars())
            {
                throw new ArgumentException("The argument must be a single web URL and cannot contain path characters.", nameof(leafUrl));
            }

            parentWeb.EnsureProperties(w => w.ServerRelativeUrl);

            var serverRelativeUrl = UrlUtility.Combine(parentWeb.ServerRelativeUrl, leafUrl);
            var webs = parentWeb.Webs;
            // NOTE: Predicate does not take into account a required case-insensitive comparison
            //var results = parentWeb.Context.LoadQuery<Web>(webs.Where(item => item.ServerRelativeUrl == serverRelativeUrl));
            parentWeb.Context.Load(webs, wc => wc.Include(w => w.ServerRelativeUrl));
            parentWeb.Context.ExecuteQueryRetry();
            var exists = webs.Any(item => string.Equals(item.ServerRelativeUrl, serverRelativeUrl, StringComparison.OrdinalIgnoreCase));
            return exists;
        }

        /// <summary>
        /// Determines if a Web (site) exists at the specified full URL, either accessible or that returns an access error.
        /// </summary>
        /// <param name="context">Existing context, used to provide credentials.</param>
        /// <param name="webFullUrl">Full URL of the site to check.</param>
        /// <returns>true if the Web (site) exists; otherwise false</returns>
        public static bool WebExistsFullUrl(this ClientRuntimeContext context, string webFullUrl)
        {
            bool exists = false;
            try
            {
                using (ClientContext testContext = context.Clone(webFullUrl))
                {
                    testContext.Load(testContext.Web, w => w.Title);
                    testContext.ExecuteQueryRetry();
                    exists = true;
                }
            }
            catch (Exception ex)
            {
                if (IsUnableToAccessSiteException(ex) || IsCannotGetSiteException(ex))
                {
                    // Site exists, but you don't have access .. not sure if this is really valid
                    // (I guess if checking if URL is already taken, e.g. want to create a new site
                    // then this makes sense).
                    exists = true;
                }
            }
            return exists;
        }

        /// <summary>
        /// Determines if a web exists by title.
        /// </summary>
        /// <param name="title">Title of the web to check.</param>
        /// <param name="parentWeb">Parent web to check under.</param>
        /// <returns>True if a web with the given title exists.</returns>
        public static bool WebExistsByTitle(this Web parentWeb, string title)
        {
            bool exists = false;

            parentWeb.EnsureProperty(p => p.Webs);

            var subWeb = (from w in parentWeb.Webs where w.Title == title select w).SingleOrDefault();
            if (subWeb != null)
            {
                exists = true;
            }
            return exists;
        }

        /// <summary>
        /// Checks if the current web is a sub site or not
        /// </summary>
        /// <param name="web">Web to check</param>
        /// <returns>True is sub site, false otherwise</returns>
        public static bool IsSubSite(this Web web)
        {
            if (web == null) throw new ArgumentNullException(nameof(web));

            var site = (web.Context as ClientContext).Site;
            var rootWeb = site.EnsureProperty(s => s.RootWeb);

            web.EnsureProperty(w => w.Id);
            rootWeb.EnsureProperty(w => w.Id);

            if (rootWeb.Id != web.Id)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Checks if the current web is a publishing site or not
        /// </summary>
        /// <param name="web">Web to check</param>
        /// <returns>True is publishing site, false otherwise</returns>
        public static bool IsPublishingWeb(this Web web)
        {
            var featureActivated = GetPropertyBagValueInternal(web, "__PublishingFeatureActivated");

            return featureActivated != null && bool.Parse(featureActivated.ToString());
        }


        /// <summary>
        /// Detects if the site in question has no script enabled or not. Detection is done by verifying if the AddAndCustomizePages permission is missing.
        ///
        /// See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f
        /// for the effects of NoScript
        ///
        /// </summary>
        /// <param name="site">site to verify</param>
        /// <returns>True if noscript, false otherwise</returns>
        public static bool IsNoScriptSite(this Site site)
        {
            return site.RootWeb.IsNoScriptSite();
        }

        /// <summary>
        /// Detects if the site in question has no script enabled or not. Detection is done by verifying if the AddAndCustomizePages permission is missing.
        ///
        /// See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f
        /// for the effects of NoScript
        ///
        /// </summary>
        /// <param name="web">Web to verify</param>
        /// <returns>True if noscript, false otherwise</returns>
        public static bool IsNoScriptSite(this Web web)
        {
#if !SP2013 && !SP2016
            web.EnsureProperties(w => w.EffectiveBasePermissions);

            // Definition of no-script is not having the AddAndCustomizePages permission
            if (!web.EffectiveBasePermissions.Has(PermissionKind.AddAndCustomizePages))
            {
                return true;
            }

            return false;
#else
            return false;
#endif
        }

        private static bool IsCannotGetSiteException(Exception ex)
        {
            if (ex is ServerException)
            {
                if (((ServerException)ex).ServerErrorCode == -1 && ((ServerException)ex).ServerErrorTypeName.Equals("Microsoft.Online.SharePoint.Common.SpoNoSiteException", StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private static bool IsUnableToAccessSiteException(Exception ex)
        {
            if (ex is ServerException)
            {
                if (((ServerException)ex).ServerErrorCode == -2147024809 && ((ServerException)ex).ServerErrorTypeName.Equals("System.ArgumentException", StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Apps and sandbox solutions

        /// <summary>
        /// Returns all app instances
        /// </summary>
        /// <param name="web">The site to process</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>all app instances</returns>
        public static ClientObjectList<AppInstance> GetAppInstances(this Web web, params Expression<Func<AppInstance, object>>[] expressions)
        {
            var instances = AppCatalog.GetAppInstances(web.Context, web);
            if (expressions != null && expressions.Any())
            {
                web.Context.Load(instances, i => i.IncludeWithDefaultProperties(expressions));
            }
            else
            {
                web.Context.Load(instances);
            }

            web.Context.Load(instances);
            web.Context.ExecuteQueryRetry();

            return instances;
        }

        /// <summary>
        /// Removes the app instance with the specified title.
        /// </summary>
        /// <param name="web">Web to remove the app instance from</param>
        /// <param name="appTitle">Title of the app instance to remove</param>
        /// <returns>true if the the app instance was removed; false if it does not exist</returns>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static bool RemoveAppInstanceByTitle(this Web web, string appTitle)
        {
            // Removes the association between the App and the Web
            bool removed = false;
            var instances = AppCatalog.GetAppInstances(web.Context, web);
            web.Context.Load(instances);
            web.Context.ExecuteQueryRetry();
            foreach (var app in instances)
            {
                if (string.Equals(app.Title, appTitle, StringComparison.OrdinalIgnoreCase))
                {
                    removed = true;
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_RemoveAppInstance, appTitle, app.Id);
                    app.Uninstall();
                    web.Context.ExecuteQueryRetry();
                }
            }
            if (!removed)
            {
                Log.Debug(Constants.LOGGING_SOURCE, "Requested to remove app '{0}', but no instances found; nothing to remove.", appTitle);
            }
            return removed;
        }

        /// <summary>
        /// Uploads and installs a sandbox solution package (.WSP) file, replacing existing solution if necessary.
        /// </summary>
        /// <param name="site">Site collection to install to</param>
        /// <param name="packageGuid">ID of the solution, from the solution manifest (required for the remove step)</param>
        /// <param name="sourceFilePath">Path to the sandbox solution package (.WSP) file</param>
        /// <param name="majorVersion">Optional major version of the solution, defaults to 1</param>
        /// <param name="minorVersion">Optional minor version of the solution, defaults to 0</param>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static void InstallSolution(this Site site, Guid packageGuid, string sourceFilePath, int majorVersion = 1, int minorVersion = 0)
        {
            string fileName = Path.GetFileName(sourceFilePath);
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_InstallSolution, fileName, site.Context.Url);

            var rootWeb = site.RootWeb;
            var sourceFileName = Path.GetFileName(sourceFilePath);

            var rootFolder = rootWeb.RootFolder;
            rootWeb.Context.Load(rootFolder, f => f.ServerRelativeUrl);
            rootWeb.Context.ExecuteQueryRetry();

            rootFolder.UploadFile(sourceFileName, sourceFilePath, true);

            var packageInfo = new DesignPackageInfo()
            {
                PackageName = fileName,
                PackageGuid = packageGuid,
                MajorVersion = majorVersion,
                MinorVersion = minorVersion,
            };

            Log.Debug(Constants.LOGGING_SOURCE, "Uninstalling package '{0}'", packageInfo.PackageName);
            UninstallSolution(site, packageGuid, fileName, majorVersion, minorVersion);
            site.Context.ExecuteQueryRetry();


            var packageServerRelativeUrl = UrlUtility.Combine(rootWeb.RootFolder.ServerRelativeUrl, fileName);
            Log.Debug(Constants.LOGGING_SOURCE, "Installing package '{0}'", packageInfo.PackageName);

            // NOTE: The lines below (in OfficeDev PnP) wipe/clear all items in the composed looks aka design catalog (_catalogs/design, list template 124).
            // The solution package should be loaded into the solutions catalog (_catalogs/solutions, list template 121).

            Publishing.DesignPackage.Install(site.Context, site, packageInfo, packageServerRelativeUrl);
            site.Context.ExecuteQueryRetry();

            // Remove package from rootfolder
            var uploadedSolutionFile = rootFolder.Files.GetByUrl(fileName);
            uploadedSolutionFile.DeleteObject();
            site.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Uninstalls a sandbox solution package (.WSP) file
        /// </summary>
        /// <param name="site">Site collection to install to</param>
        /// <param name="packageGuid">ID of the solution, from the solution manifest</param>
        /// <param name="fileName">filename of the WSP file to uninstall</param>
        /// <param name="majorVersion">Optional major version of the solution, defaults to 1</param>
        /// <param name="minorVersion">Optional minor version of the solution, defaults to 0</param>
        public static void UninstallSolution(this Site site, Guid packageGuid, string fileName, int majorVersion = 1, int minorVersion = 0)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_UninstallSolution, packageGuid);

            var rootWeb = site.RootWeb;
            var solutionGallery = rootWeb.GetCatalog((int)ListTemplateType.SolutionCatalog);

            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = $@"<View>  
                        <Query> 
                           <Where><Eq><FieldRef Name='SolutionId' /><Value Type='Guid'>{packageGuid}</Value></Eq></Where> 
                        </Query> 
                         <ViewFields><FieldRef Name='ID' /><FieldRef Name='FileLeafRef' /></ViewFields> 
                  </View>";

            var solutions = solutionGallery.GetItems(camlQuery);
            site.Context.Load(solutions);
            site.Context.ExecuteQueryRetry();

            if (solutions.AreItemsAvailable && solutions.Count > 0)
            {
                var packageItem = solutions.FirstOrDefault();
                var packageInfo = new DesignPackageInfo()
                {
                    PackageGuid = packageGuid,
                    PackageName = fileName,
                    MajorVersion = majorVersion,
                    MinorVersion = minorVersion
                };

                Publishing.DesignPackage.UnInstall(site.Context, site, packageInfo);
                site.Context.ExecuteQueryRetry();
            }
        }

        #endregion

        #region Site retrieval via search
        /// <summary>
        /// Returns all my site site collections
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <returns>All my site site collections</returns>
        [SuppressMessage("Microsoft.Usage", "CA2241:Provide correct arguments to formatting methods",
            Justification = "Search Query code")]
        public static List<SiteEntity> MySiteSearch(this Web web)
        {
            const string keywordQuery = "contentclass:\"STS_Site\" AND WebTemplate:SPSPERS";
            return web.SiteSearch(keywordQuery);
        }

        /// <summary>
        /// Returns all site collections that are indexed. In MT the search center, mysite host and contenttype hub are defined as non indexable by default and thus
        /// are not returned
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <returns>All site collections</returns>
        public static List<SiteEntity> SiteSearch(this Web web)
        {
            return web.SiteSearch(string.Empty);
        }

        /// <summary>
        /// Returns the site collections that comply with the passed keyword query
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="keywordQueryValue">Keyword query</param>
        /// <param name="trimDuplicates">Indicates if duplicates should be trimmed or not</param>
        /// <returns>All found site collections</returns>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static List<SiteEntity> SiteSearch(this Web web, string keywordQueryValue, bool trimDuplicates = false)
        {
            try
            {
                Log.Debug(Constants.LOGGING_SOURCE, "Site search '{0}'", keywordQueryValue);

                List<SiteEntity> sites = new List<SiteEntity>();

                KeywordQuery keywordQuery = new KeywordQuery(web.Context);
                keywordQuery.TrimDuplicates = trimDuplicates;

                if (keywordQueryValue.Length == 0)
                {

                    keywordQueryValue = "contentclass:\"STS_Site\"";

                }

                //int startRow = 0;
                int totalRows = 0;

                totalRows = web.ProcessQuery(keywordQueryValue, sites, keywordQuery);


                if (totalRows > 0)
                {
                    while (totalRows > 0)
                    {
                        totalRows = web.ProcessQuery(keywordQueryValue + " AND IndexDocId >" + sites.Last().IndexDocId, sites, keywordQuery);// From the second Query get the next set (rowlimit) of search result based on IndexDocId
                    }
                }

                return sites;
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_SiteSearchUnhandledException, ex.Message);
                // rethrow does lose one line of stack trace, but we want to log the error at the component boundary
                throw;
            }
        }

        /// <summary>
        /// Returns all site collection that start with the provided URL
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="siteUrl">Base URL for which sites can be returned</param>
        /// <returns>All found site collections</returns>
        public static List<SiteEntity> SiteSearchScopedByUrl(this Web web, string siteUrl)
        {
            string keywordQuery = $"contentclass:\"STS_Site\" AND site:{siteUrl}";
            return web.SiteSearch(keywordQuery);
        }

        /// <summary>
        /// Returns all site collection that match with the provided title
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="siteTitle">Title of the site to search for</param>
        /// <returns>All found site collections</returns>
        public static List<SiteEntity> SiteSearchScopedByTitle(this Web web, string siteTitle)
        {
            string keywordQuery = $"contentclass:\"STS_Site\" AND Title:{siteTitle}";
            return web.SiteSearch(keywordQuery);
        }

        // private methods
        /// <summary>
        /// Runs a query
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="keywordQueryValue">keyword query </param>
        /// <param name="sites">sites variable that hold the resulting sites</param>
        /// <param name="keywordQuery">KeywordQuery object</param>

        /// <returns>Total number of rows for the query</returns>
        private static int ProcessQuery(this Web web, string keywordQueryValue, List<SiteEntity> sites, KeywordQuery keywordQuery)
        {
            int totalRows = 0;

            keywordQuery.QueryText = keywordQueryValue;
            keywordQuery.RowLimit = 500;
            // keywordQuery.StartRow = startRow;
            keywordQuery.SelectProperties.Add("Title");
            keywordQuery.SelectProperties.Add("SPSiteUrl");
            keywordQuery.SelectProperties.Add("Description");
            keywordQuery.SelectProperties.Add("WebTemplate");
            keywordQuery.SelectProperties.Add("IndexDocId"); // Change : Include IndexDocId property to get the IndexDocId for paging
            keywordQuery.SortList.Add("IndexDocId", SortDirection.Ascending); // Change : Sort by IndexDocId
            SearchExecutor searchExec = new SearchExecutor(web.Context);

            // Important to avoid trimming "similar" site collections
            keywordQuery.TrimDuplicates = false;

            ClientResult<ResultTableCollection> results = searchExec.ExecuteQuery(keywordQuery);
            web.Context.ExecuteQueryRetry();

            if (results != null)
            {
                if (results.Value[0].RowCount > 0)
                {
                    totalRows = results.Value[0].TotalRows;

                    foreach (var row in results.Value[0].ResultRows)
                    {
                        sites.Add(new SiteEntity
                        {
                            Title = row["Title"] != null ? row["Title"].ToString() : "",
                            Url = row["SPSiteUrl"] != null ? row["SPSiteUrl"].ToString() : "",
                            Description = row["Description"] != null ? row["Description"].ToString() : "",
                            Template = row["WebTemplate"] != null ? row["WebTemplate"].ToString() : "",
                            IndexDocId = row["DocId"] != null ? double.Parse(row["DocId"].ToString()) : 0, // Change : Include IndexDocId in the sites List
                        });
                    }
                }
            }

            return totalRows;
        }
        #endregion

        #region Web (site) Property Bag Modifiers

        /// <summary>
        /// Sets a key/value pair in the web property bag
        /// </summary>
        /// <param name="web">Web that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">Integer value for the property bag entry</param>
        public static void SetPropertyBagValue(this Web web, string key, int value)
        {
            SetPropertyBagValueInternal(web, key, value);
        }


        /// <summary>
        /// Sets a key/value pair in the web property bag
        /// </summary>
        /// <param name="web">Web that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">String value for the property bag entry</param>
        public static void SetPropertyBagValue(this Web web, string key, string value)
        {
            SetPropertyBagValueInternal(web, key, value);
        }

        /// <summary>
        /// Sets a key/value pair in the web property bag
        /// </summary>
        /// <param name="web">Web that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">Datetime value for the property bag entry</param>
        public static void SetPropertyBagValue(this Web web, string key, DateTime value)
        {
            SetPropertyBagValueInternal(web, key, value);
        }

        /// <summary>
        /// Sets a key/value pair in the web property bag
        /// </summary>
        /// <param name="web">Web that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">Value for the property bag entry</param>
        private static void SetPropertyBagValueInternal(Web web, string key, object value)
        {
            web.AllProperties.ClearObjectData();

            var props = web.AllProperties;

            // Get the value, if the web properties are already loaded
            if (props.FieldValues.Count > 0)
            {
                props[key] = value;
            }
            else
            {
                // Load the web properties
                web.Context.Load(props);
                web.Context.ExecuteQueryRetry();

                props[key] = value;
            }

            web.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Removes a property bag value from the property bag
        /// </summary>
        /// <param name="web">The site to process</param>
        /// <param name="key">The key to remove</param>
        public static void RemovePropertyBagValue(this Web web, string key)
        {
            RemovePropertyBagValueInternal(web, key, true);
        }

        /// <summary>
        /// Removes a property bag value
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="key">They key to remove</param>
        /// <param name="checkIndexed"></param>
        private static void RemovePropertyBagValueInternal(Web web, string key, bool checkIndexed)
        {
            // In order to remove a property from the property bag, remove it both from the AllProperties collection by setting it to null
            // -and- by removing it from the FieldValues collection. Bug in CSOM?
            web.AllProperties[key] = null;
            web.AllProperties.FieldValues.Remove(key);

            web.Update();

            web.Context.ExecuteQueryRetry();
            if (checkIndexed)
                RemoveIndexedPropertyBagKey(web, key); // Will only remove it if it exists as an indexed property
        }

        /// <summary>
        /// Get int typed property bag value. If does not contain, returns default value.
        /// </summary>
        /// <param name="web">Web to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <param name="defaultValue"></param>
        /// <returns>Value of the property bag entry as integer</returns>
        public static int? GetPropertyBagValueInt(this Web web, string key, int defaultValue)
        {
            object value = GetPropertyBagValueInternal(web, key);
            if (value != null)
            {
                return (int)value;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Get DateTime typed property bag value. If does not contain, returns default value.
        /// </summary>
        /// <param name="web">Web to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <param name="defaultValue"></param>
        /// <returns>Value of the property bag entry as integer</returns>
        public static DateTime? GetPropertyBagValueDateTime(this Web web, string key, DateTime defaultValue)
        {
            object value = GetPropertyBagValueInternal(web, key);
            if (value != null)
            {
                return (DateTime)value;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Get string typed property bag value. If does not contain, returns given default value.
        /// </summary>
        /// <param name="web">Web to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <param name="defaultValue"></param>
        /// <returns>Value of the property bag entry as string</returns>
        public static string GetPropertyBagValueString(this Web web, string key, string defaultValue)
        {
            object value = GetPropertyBagValueInternal(web, key);
            if (value != null)
            {
                return (string)value;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Type independent implementation of the property getter.
        /// </summary>
        /// <param name="web">Web to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <returns>Value of the property bag entry</returns>
        private static object GetPropertyBagValueInternal(Web web, string key)
        {
            web.AllProperties.ClearObjectData();

            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();
            if (props.FieldValues.ContainsKey(key))
            {
                return props.FieldValues[key];
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// Checks if the given property bag entry exists
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="key">Key of the property bag entry to check</param>
        /// <returns>True if the entry exists, false otherwise</returns>
        public static bool PropertyBagContainsKey(this Web web, string key)
        {
            web.AllProperties.ClearObjectData();

            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();
            if (props.FieldValues.ContainsKey(key))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Used to convert the list of property keys is required format for listing keys to be index
        /// </summary>
        /// <param name="keys">list of keys to set to be searchable</param>
        /// <returns>string formatted list of keys in proper format</returns>
        private static string GetEncodedValueForSearchIndexProperty(IEnumerable<string> keys)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (string current in keys)
            {
                stringBuilder.Append(Convert.ToBase64String(Encoding.Unicode.GetBytes(current)));
                stringBuilder.Append('|');
            }
            return stringBuilder.ToString();
        }

        /// <summary>
        /// Returns all keys in the property bag that have been marked for indexing
        /// </summary>
        /// <param name="web">The site to process</param>
        /// <returns>all indexed property bag keys</returns>
        public static IEnumerable<string> GetIndexedPropertyBagKeys(this Web web)
        {
            var keys = new List<string>();

            if (web.PropertyBagContainsKey(INDEXED_PROPERTY_KEY))
            {
                foreach (var key in web.GetPropertyBagValueString(INDEXED_PROPERTY_KEY, "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var bytes = Convert.FromBase64String(key);
                    keys.Add(Encoding.Unicode.GetString(bytes));
                }
            }

            return keys;
        }

        /// <summary>
        /// Marks a property bag key for indexing
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="key">The key to mark for indexing</param>
        /// <returns>Returns True if succeeded</returns>
        public static bool AddIndexedPropertyBagKey(this Web web, string key)
        {
            var result = false;
            var keys = GetIndexedPropertyBagKeys(web).ToList();
            if (!keys.Contains(key))
            {
                keys.Add(key);
                web.SetPropertyBagValue(INDEXED_PROPERTY_KEY, GetEncodedValueForSearchIndexProperty(keys));
                result = true;
            }
            return result;
        }

        /// <summary>
        /// Unmarks a property bag key for indexing
        /// </summary>
        /// <param name="web">The site to process</param>
        /// <param name="key">The key to unmark for indexed. Case-sensitive</param>
        /// <returns>Returns True if succeeded</returns>
        public static bool RemoveIndexedPropertyBagKey(this Web web, string key)
        {
            var result = false;
            var keys = GetIndexedPropertyBagKeys(web).ToList();
            if (key.Contains(key))
            {
                keys.Remove(key);
                if (keys.Any())
                {
                    web.SetPropertyBagValue(INDEXED_PROPERTY_KEY, GetEncodedValueForSearchIndexProperty(keys));
                }
                else
                {
                    RemovePropertyBagValueInternal(web, INDEXED_PROPERTY_KEY, false);
                }
                result = true;
            }
            return result;
        }

        #endregion

        #region Search

        /// <summary>
        /// Queues a web for a full crawl the next incremental/continous crawl
        /// </summary>
        /// <param name="web">Site to be processed</param>
        public static void ReIndexWeb(this Web web)
        {
            if (web.IsNoScriptSite())
            {
                // Update individual lists instead, as web bag is no (longer) accessible
                var context = web.Context;
                context.Load(web.Lists);
                context.ExecuteQueryRetry();
                foreach (var list in web.Lists)
                {
                    list.ReIndexList();
                }
            }
            else
            {
                int searchversion = 0;
                if (web.PropertyBagContainsKey("vti_searchversion"))
                {
                    searchversion = (int)web.GetPropertyBagValueInt("vti_searchversion", 0);
                }
                web.SetPropertyBagValue("vti_searchversion", searchversion + 1);
            }
        }
        #endregion

        #region Events


        /// <summary>
        /// Registers a remote event receiver
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="name">The name of the event receiver (needs to be unique among the event receivers registered on this list)</param>
        /// <param name="url">The URL of the remote WCF service that handles the event</param>
        /// <param name="eventReceiverType"></param>
        /// <param name="synchronization"></param>
        /// <param name="force">If True any event already registered with the same name will be removed first.</param>
        /// <returns>Returns an EventReceiverDefinition if succeeded. Returns null if failed.</returns>
        public static EventReceiverDefinition AddRemoteEventReceiver(this Web web, string name, string url, EventReceiverType eventReceiverType, EventReceiverSynchronization synchronization, bool force)
        {
            return web.AddRemoteEventReceiver(name, url, eventReceiverType, synchronization, 1000, force);
        }

        /// <summary>
        /// Registers a remote event receiver
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="name">The name of the event receiver (needs to be unique among the event receivers registered on this list)</param>
        /// <param name="url">The URL of the remote WCF service that handles the event</param>
        /// <param name="eventReceiverType">The type of event for the event receiver.</param>
        /// <param name="synchronization">An enumeration that specifies the synchronization state for the event receiver.</param>
        /// <param name="sequenceNumber">An integer that represents the relative sequence of the event.</param>
        /// <param name="force">If True any event already registered with the same name will be removed first.</param>
        /// <returns>Returns an EventReceiverDefinition if succeeded. Returns null if failed.</returns>
        public static EventReceiverDefinition AddRemoteEventReceiver(this Web web, string name, string url, EventReceiverType eventReceiverType, EventReceiverSynchronization synchronization, int sequenceNumber, bool force)
        {
            var query = from receiver
                   in web.EventReceivers
                        where receiver.ReceiverName == name
                        select receiver;
            var receivers = web.Context.LoadQuery(query);
            web.Context.ExecuteQueryRetry();

            var receiverExists = receivers.Any();
            if (receiverExists && force)
            {
                var receiver = receivers.FirstOrDefault();
                receiver.DeleteObject();
                web.Context.ExecuteQueryRetry();
                receiverExists = false;
            }
            EventReceiverDefinition def = null;

            if (!receiverExists)
            {
                EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();
                receiver.EventType = eventReceiverType;
                receiver.ReceiverUrl = url;
                receiver.ReceiverName = name;
                receiver.SequenceNumber = sequenceNumber;
                receiver.Synchronization = synchronization;
                def = web.EventReceivers.Add(receiver);
                web.Context.Load(def);
                web.Context.ExecuteQueryRetry();
            }
            return def;
        }

        /// <summary>
        /// Returns an event receiver definition
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="id">The id of event receiver</param>
        /// <returns>Returns an EventReceiverDefinition if succeeded. Returns null if failed.</returns>
        public static EventReceiverDefinition GetEventReceiverById(this Web web, Guid id)
        {
            IEnumerable<EventReceiverDefinition> receivers = null;
            var query = from receiver
                        in web.EventReceivers
                        where receiver.ReceiverId == id
                        select receiver;

            receivers = web.Context.LoadQuery(query);
            web.Context.ExecuteQueryRetry();
            if (receivers.Any())
            {
                return receivers.FirstOrDefault();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns an event receiver definition
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="name">The name of the receiver</param>
        /// <returns>Returns an EventReceiverDefinition if succeeded. Returns null if failed.</returns>
        public static EventReceiverDefinition GetEventReceiverByName(this Web web, string name)
        {
            IEnumerable<EventReceiverDefinition> receivers = null;
            var query = from receiver
                        in web.EventReceivers
                        where receiver.ReceiverName == name
                        select receiver;

            receivers = web.Context.LoadQuery(query);
            web.Context.ExecuteQueryRetry();
            if (receivers.Any())
            {
                return receivers.FirstOrDefault();
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region Localization
#if !ONPREMISES
        /// <summary>
        /// Can be used to set translations for different cultures.
        /// </summary>
        /// <example>
        ///     web.SetLocalizationForSiteLabels("fi-fi", "Name of the site in Finnish", "Description in Finnish");
        /// </example>
        /// <see href="http://blogs.msdn.com/b/vesku/archive/2014/03/20/office365-multilingual-content-types-site-columns-and-site-other-elements.aspx"/>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="cultureName">Culture name like en-us or fi-fi</param>
        /// <param name="titleResource">Localized Title string</param>
        /// <param name="descriptionResource">Localized Description string</param>
        public static void SetLocalizationLabels(this Web web, string cultureName, string titleResource, string descriptionResource)
        {
            web.EnsureProperties(w => w.TitleResource);

            // Set translations for the culture
            web.TitleResource.SetValueForUICulture(cultureName, titleResource);
            web.DescriptionResource.SetValueForUICulture(cultureName, descriptionResource);
            web.Update();
            web.Context.ExecuteQueryRetry();
        }
#endif
        #endregion

        #region TemplateHandling

        /// <summary>
        /// Can be used to apply custom remote provisioning template on top of existing site.
        /// </summary>
        /// <param name="web">web to apply remote template</param>
        /// <param name="template">ProvisioningTemplate with the settings to be applied</param>
        /// <param name="applyingInformation">Specified additional settings and or properties</param>
        public static void ApplyProvisioningTemplate(this Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation = null)
        {
            // Call actual handler
            new SiteToTemplateConversion().ApplyRemoteTemplate(web, template, applyingInformation);
        }

        /// <summary>
        /// Can be used to extract custom provisioning template from existing site. The extracted template
        /// will be compared with the default base template.
        /// </summary>
        /// <param name="web">Web to get template from</param>
        /// <returns>ProvisioningTemplate object with generated values from existing site</returns>
        public static ProvisioningTemplate GetProvisioningTemplate(this Web web)
        {
            ProvisioningTemplateCreationInformation creationInfo = new ProvisioningTemplateCreationInformation(web);

            return new SiteToTemplateConversion().GetRemoteTemplate(web, creationInfo);
        }

        /// <summary>
        /// Can be used to extract custom provisioning template from existing site. The extracted template
        /// will be compared with the default base template.
        /// </summary>
        /// <param name="web">Web to get template from</param>
        /// <param name="creationInfo">Specifies additional settings and/or properties</param>
        /// <returns>ProvisioningTemplate object with generated values from existing site</returns>
        public static ProvisioningTemplate GetProvisioningTemplate(this Web web, ProvisioningTemplateCreationInformation creationInfo)
        {
            return new SiteToTemplateConversion().GetRemoteTemplate(web, creationInfo);
        }

        #endregion

        #region Output Cache

        /// <summary>
        /// Sets output cache on publishing web. The settings can be maintained from UI by visiting URL /_layouts/15/sitecachesettings.aspx
        /// </summary>
        /// <param name="web">SharePoint web</param>
        /// <param name="enableOutputCache">Specify true to enable output cache. False otherwise.</param>
        /// <param name="anonymousCacheProfileId">Applies for anonymous users access for a site in Site Collection. Id of the profile specified in "Cache Profiles" list.</param>
        /// <param name="authenticatedCacheProfileId">Applies for authenticated users access for a site in the Site Collection. Id of the profile specified in "Cache Profiles" list.</param>
        /// <param name="debugCacheInformation">Specify true to enable the display of additional cache information on pages in this site collection. False otherwise.</param>
        public static void SetPageOutputCache(this Web web, bool enableOutputCache, int anonymousCacheProfileId, int authenticatedCacheProfileId, bool debugCacheInformation)
        {
            const string cacheProfileUrl = "Cache Profiles/{0}_.000";

            string publishingWebValue = web.GetPropertyBagValueString("__PublishingFeatureActivated", string.Empty);
            if (string.IsNullOrEmpty(publishingWebValue))
            {
                throw new Exception("Page output cache can be set only on publishing sites.");
            }

            web.SetPropertyBagValue("EnableCache", enableOutputCache.ToString());
            web.SetPropertyBagValue("AnonymousPageCacheProfileUrl", string.Format(cacheProfileUrl, anonymousCacheProfileId));
            web.SetPropertyBagValue("AuthenticatedPageCacheProfileUrl", string.Format(cacheProfileUrl, authenticatedCacheProfileId));
            web.SetPropertyBagValue("EnableDebuggingOutput", debugCacheInformation.ToString());
        }

        #endregion

        #region Request Access
        /// <summary>
        /// Disables the request access on the web.
        /// </summary>
        /// <param name="web">The web to disable request access.</param>
        public static void DisableRequestAccess(this Web web)
        {
            web.RequestAccessEmail = string.Empty;
            web.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Enables request access for the specified e-mail addresses.
        /// </summary>
        /// <param name="web">The web to enable request access.</param>
        /// <param name="emails">The e-mail addresses to send access requests to.</param>
        public static void EnableRequestAccess(this Web web, params string[] emails)
        {
            web.EnableRequestAccess(emails.AsEnumerable());
        }

        /// <summary>
        /// Enables request access for the specified e-mail addresses.
        /// </summary>
        /// <param name="web">The web to enable request access.</param>
        /// <param name="emails">The e-mail addresses to send access requests to.</param>
        public static void EnableRequestAccess(this Web web, IEnumerable<string> emails)
        {
            // keep them unique, but keep order
            var uniqueEmails = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            var sb = new StringBuilder();
            var skippedEmails = new List<string>();

            foreach (string email in emails)
            {
                if (uniqueEmails.Contains(email) || string.IsNullOrWhiteSpace(email))
                    continue;

                var value = (sb.Length > 0 ? ";" : "") + email;

                // max 255 chars
                if (sb.Length + value.Length <= byte.MaxValue)
                {
                    sb.Append(value);
                    uniqueEmails.Add(email);
                }
                else
                {
                    skippedEmails.Add(email);
                }
            }

            if (skippedEmails.Count > 0)
                Log.Warning(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_RequestAccessEmailLimitExceeded, string.Join(", ", skippedEmails));

            web.RequestAccessEmail = sb.ToString();
            web.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Gets the request access e-mail addresses of the web.
        /// </summary>
        /// <param name="web">The web to get the request access e-mail addresses from.</param>
        /// <returns>The request access e-mail addresses of the web.</returns>
        public static IEnumerable<string> GetRequestAccessEmails(this Web web)
        {
            var emails = new List<string>();

            web.EnsureProperty(w => w.RequestAccessEmail);

            if (!string.IsNullOrWhiteSpace(web.RequestAccessEmail))
            {
                foreach (string email in web.RequestAccessEmail.Split(';'))
                    emails.Add(email.Trim());
            }

            return emails;
        }
        #endregion

        /// <summary>
        /// Gets the name part of the URL of the Server Relative URL of the Web.
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <returns>A string that contains the name part of the Server Relative URL (the last part of the URL) of a web.</returns>
        public static string GetName(this Web web)
        {
            web.Context.Load(web, w => w.ParentWeb.ServerRelativeUrl);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            string webName;
            string parentWebUrl = null;

            //web.ParentWeb.ServerObjectIsNull will be null if a parent web exists.
            //ClientObjectExtensions.ServerObjectIsNull() seems to have a problem when
            //ClientObject.ServerObjectIsNull == null
            //ServerObjectIsNull is then undefined but ClientObjectExtensions.ServerObjectIsNull()
            //incorrectly returns true.
            if (web.ParentWeb.ServerObjectIsNull == null || !web.ParentWeb.ServerObjectIsNull.Value)
            {
                parentWebUrl = web.ParentWeb.ServerRelativeUrl;
            }

            if (parentWebUrl == null)
            {
                webName = string.Empty;
            }
            else
            {
                webName = UrlUtility.ConvertToServiceRelUrl(web.ServerRelativeUrl, parentWebUrl);
            }
            return webName;
        }

#if !SP2013 && !SP2016
        #region ClientSide Package Deployment
        /// <summary>
        /// Gets the Uri for the tenant's app catalog site (if that one has already been created)
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <returns>The Uri holding the app catalog site URL</returns>
        public static Uri GetAppCatalog(this Web web)
        {
            var tenantSettings = TenantSettings.GetCurrent(web.Context);
            tenantSettings.EnsureProperties(s => s.CorporateCatalogUrl);
            if(!string.IsNullOrEmpty(tenantSettings.CorporateCatalogUrl))
            {
                return new Uri(tenantSettings.CorporateCatalogUrl);
            }
            return null;
        }

        /// <summary>
        /// Adds a package to the tenants app catalog and by default deploys it if the package is a client side package (sppkg)
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="spPkgName">Name of the package to upload (e.g. demo.sppkg) </param>
        /// <param name="spPkgPath">Path on the filesystem where this package is stored</param>
        /// <param name="autoDeploy">Automatically deploy the package, only applies to client side packages (sppkg)</param>
        /// <param name="overwrite">Overwrite the package if it was already listed in the app catalog</param>
        /// <returns>The ListItem of the added package row</returns>
        public static ListItem DeployApplicationPackageToAppCatalog(this Web web, string spPkgName, string spPkgPath, bool autoDeploy = true, bool overwrite = true)
        {
            var appCatalogSite = web.GetAppCatalog();
            if (appCatalogSite == null)
            {
                throw new ArgumentException("No app catalog site found, please ensure the site exists or specify the site as parameter. Note that the app catalog site is retrieved via search, so take in account the indexing time.");
            }

            return DeployApplicationPackageToAppCatalogImplementation(web, appCatalogSite.ToString(), spPkgName, spPkgPath, autoDeploy, false, overwrite);
        }

        /// <summary>
        /// Adds a package to the tenants app catalog and by default deploys it if the package is a client side package (sppkg)
        /// </summary>
        /// <param name="web">Tenant to operate against</param>
        /// <param name="spPkgName">Name of the package to upload (e.g. demo.sppkg) </param>
        /// <param name="spPkgPath">Path on the filesystem where this package is stored</param>
        /// <param name="autoDeploy">Automatically deploy the package, only applies to client side packages (sppkg)</param>
        /// <param name="skipFeatureDeployment">Skip the feature deployment step, allows for a one-time central deployment of your solution</param>
        /// <param name="overwrite">Overwrite the package if it was already listed in the app catalog</param>
        /// <returns>The ListItem of the added package row</returns>
        public static ListItem DeployApplicationPackageToAppCatalog(this Web web, string spPkgName, string spPkgPath, bool autoDeploy = true, bool skipFeatureDeployment = true, bool overwrite = true)
        {
            var appCatalogSite = web.GetAppCatalog();
            if (appCatalogSite == null)
            {
                throw new ArgumentException("No app catalog site found, please ensure the site exists or specify the site as parameter. Note that the app catalog site is retrieved via search, so take in account the indexing time.");
            }

            return DeployApplicationPackageToAppCatalogImplementation(web, appCatalogSite.ToString(), spPkgName, spPkgPath, autoDeploy, skipFeatureDeployment, overwrite);
        }


        private static ListItem DeployApplicationPackageToAppCatalogImplementation(this Web web, string appCatalogSiteUrl, string spPkgName, string spPkgPath, bool autoDeploy, bool skipFeatureDeployment, bool overwrite)
        {
            if (String.IsNullOrEmpty(appCatalogSiteUrl))
            {
                throw new ArgumentException("Please specify a app catalog site URL");
            }

            Uri catalogUri;
            if (!Uri.TryCreate(appCatalogSiteUrl, UriKind.Absolute, out catalogUri))
            {
                throw new ArgumentException("Please specify a valid app catalog site URL");
            }

            if (String.IsNullOrEmpty(spPkgName))
            {
                throw new ArgumentException("Please specify a package name");
            }

            if (String.IsNullOrEmpty(spPkgPath))
            {
                throw new ArgumentException("Please specify a package path");
            }

            using (var appCatalogContext = web.Context.Clone(catalogUri))
            {
                List catalog = appCatalogContext.Web.GetListByUrl("appcatalog");
                if (catalog == null)
                {
                    throw new Exception($"No app catalog found...did you provide a valid app catalog site?");
                }

                Folder rootFolder = catalog.RootFolder;

                // Upload package
                var sppkgFile = rootFolder.UploadFile(spPkgName, System.IO.Path.Combine(spPkgPath, spPkgName), overwrite);
                if (sppkgFile == null)
                {
                    throw new Exception($"Upload of {spPkgName} failed");
                }

                if ((autoDeploy || skipFeatureDeployment) &&
                    System.IO.Path.GetExtension(spPkgName).ToLower() == ".sppkg")
                {
                    // Trigger "deployment" by setting the IsClientSideSolutionDeployed bool to true which triggers
                    // an event receiver that will process the sppkg file and update the client side componenent manifest list
                    sppkgFile.ListItemAllFields["IsClientSideSolutionDeployed"] = autoDeploy;
                    // deal with "upgrading" solutions
                    sppkgFile.ListItemAllFields["IsClientSideSolutionCurrentVersionDeployed"] = autoDeploy;
                    // Allow for a central deployment of the solution, no need to install the solution in the individual site collections.
                    // Only works when the solution is not using feature framework to "configure" the site upon solution installation
                    sppkgFile.ListItemAllFields["SkipFeatureDeployment"] = skipFeatureDeployment;
                    sppkgFile.ListItemAllFields.Update();
                }

                appCatalogContext.Load(sppkgFile.ListItemAllFields);
                appCatalogContext.ExecuteQueryRetry();

                return sppkgFile.ListItemAllFields;
            }
        }
        #endregion
#endif

    }
}
