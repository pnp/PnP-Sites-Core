using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
#if !NETSTANDARD2_0
using System.Xml.Serialization.Configuration;
#endif
using Microsoft.Graph;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
#if !NETSTANDARD2_0
using OfficeDevPnP.Core.UPAWebService;
#endif
using OfficeDevPnP.Core.Diagnostics;
using System.Net.Http;
using CoreUtilities = OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Framework.Graph;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Framework.Graph.Model;
using Newtonsoft.Json;
#if !ONPREMISES
using OfficeDevPnP.Core.Sites;
#endif

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for tenant extension methods
    /// </summary>
    public static partial class TenantExtensions
    {
        const string SITE_STATUS_RECYCLED = "Recycled";

#if !ONPREMISES
        #region Site collection creation
        /// <summary>
        /// Adds a SiteEntity by launching site collection creation and waits for the creation to finish
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="properties">Describes the site collection to be created</param>
        /// <param name="removeFromRecycleBin">It true and site is present in recycle bin, it will be removed first from the recycle bin</param>
        /// <param name="wait">If true, processing will halt until the site collection has been created</param>
        /// <param name="timeoutFunction">An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.</param>
        /// <returns>Guid of the created site collection and Guid.Empty is the wait parameter is specified as false. Returns Guid.Empty if the wait is cancelled.</returns>
        public static Guid CreateSiteCollection(this Tenant tenant, SiteEntity properties, bool removeFromRecycleBin = false, bool wait = true, Func<TenantOperationMessage, bool> timeoutFunction = null)
        {
            if (removeFromRecycleBin)
            {
                if (tenant.CheckIfSiteExists(properties.Url, SITE_STATUS_RECYCLED))
                {
                    tenant.DeleteSiteCollectionFromRecycleBin(properties.Url);
                }
            }

            SiteCreationProperties newsite = new SiteCreationProperties();
            newsite.Url = properties.Url;
            newsite.Owner = properties.SiteOwnerLogin;
            newsite.Template = properties.Template;
            newsite.Title = properties.Title;
            newsite.StorageMaximumLevel = properties.StorageMaximumLevel;
            newsite.StorageWarningLevel = properties.StorageWarningLevel;
            newsite.TimeZoneId = properties.TimeZoneId;
            newsite.UserCodeMaximumLevel = properties.UserCodeMaximumLevel;
            newsite.UserCodeWarningLevel = properties.UserCodeWarningLevel;
            newsite.Lcid = properties.Lcid;

            SpoOperation op = tenant.CreateSite(newsite);
            tenant.Context.Load(tenant);
            tenant.Context.Load(op, i => i.IsComplete, i => i.PollingInterval);
            tenant.Context.ExecuteQueryRetry();

            // Get site guid and return. If we create the site asynchronously, return an empty guid as we cannot retrieve the site by URL yet.
            Guid siteGuid = Guid.Empty;
            if (timeoutFunction != null)
            {
                wait = true;
            }
            if (wait)
            {
                // Let's poll for site collection creation completion
                if (WaitForIsComplete(tenant, op, timeoutFunction, TenantOperationMessage.CreatingSiteCollection))
                {
                    // Restore of original flow to validate correct working in edog after fix got committed
                    if (properties.Url.ToLower().Contains("spoppe.com"))
                    {
                        siteGuid = tenant.GetSiteGuidByUrl(new Uri(properties.Url));
                    }
                    else
                    {
                        // Return site guid of created site collection
                        try
                        {
                            siteGuid = tenant.GetSiteGuidByUrl(new Uri(properties.Url));
                        }
                        catch (Exception ex)
                        {
                            // Eat all exceptions cause there's currently (December 16) an issue in the service that can make tenant API calls fail in combination with app-only usage
                            Log.Error("Temp eating exception due to issue in service (December 2016). Exception is {0}.",
                                ex.ToDetailedString());
                        }
                    }
                }
            }
            return siteGuid;
        }

        /// <summary>
        /// Launches a site collection creation and waits for the creation to finish 
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">The SPO URL</param>
        /// <param name="title">The site title</param>
        /// <param name="siteOwnerLogin">Owner account</param>
        /// <param name="template">Site template being used</param>
        /// <param name="storageMaximumLevel">Site quota in MB</param>
        /// <param name="storageWarningLevel">Site quota warning level in MB</param>
        /// <param name="timeZoneId">TimeZoneID for the site. "(UTC+01:00) Brussels, Copenhagen, Madrid, Paris" = 3 </param>
        /// <param name="userCodeMaximumLevel">The user code quota in points</param>
        /// <param name="userCodeWarningLevel">The user code quota warning level in points</param>
        /// <param name="lcid">The site locale. See http://technet.microsoft.com/en-us/library/ff463597.aspx for a complete list of Lcid's</param>
        /// <param name="removeFromRecycleBin">If true, any existing site with the same URL will be removed from the recycle bin</param>
        /// <param name="wait">Wait for the site to be created before continuing processing</param>
        /// <param name="timeoutFunction">An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.</param>
        /// <returns>Guid of the created site collection and Guid.Empty is the wait parameter is specified as false. Returns Guid.Empty if the wait is cancelled.</returns>
        public static Guid CreateSiteCollection(this Tenant tenant, string siteFullUrl, string title, string siteOwnerLogin,
                                                        string template, int storageMaximumLevel, int storageWarningLevel,
                                                        int timeZoneId, int userCodeMaximumLevel, int userCodeWarningLevel,
                                                        uint lcid, bool removeFromRecycleBin = false, bool wait = true, Func<TenantOperationMessage, bool> timeoutFunction = null)
        {
            SiteEntity siteCol = new SiteEntity()
            {
                Url = siteFullUrl,
                Title = title,
                SiteOwnerLogin = siteOwnerLogin,
                Template = template,
                StorageMaximumLevel = storageMaximumLevel,
                StorageWarningLevel = storageWarningLevel,
                TimeZoneId = timeZoneId,
                UserCodeMaximumLevel = userCodeMaximumLevel,
                UserCodeWarningLevel = userCodeWarningLevel,
                Lcid = lcid
            };
            return tenant.CreateSiteCollection(siteCol, removeFromRecycleBin, wait, timeoutFunction);
        }
        #endregion

        #region Site status checks
        /// <summary>
        /// Returns if a site collection is in a particular status. If the URL contains a sub site then returns true is the sub site exists, false if not. 
        /// Status is irrelevant for sub sites
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">Url to the site collection</param>
        /// <param name="status">Status to check (Active, Creating, Recycled)</param>
        /// <returns>True if in status, false if not in status</returns>
        public static bool CheckIfSiteExists(this Tenant tenant, string siteFullUrl, string status)
        {
            bool ret = false;
            //Get the site name
            var url = new Uri(siteFullUrl);
            var siteDomainUrl = url.GetLeftPart(UriPartial.Authority);
            int siteNameIndex = url.AbsolutePath.IndexOf('/', 1) + 1;
            var managedPath = url.AbsolutePath.Substring(0, siteNameIndex);
            var siteRelativePath = url.AbsolutePath.Substring(siteNameIndex);
            var isSiteCollection = siteRelativePath.IndexOf('/') == -1;

            //Judge whether this site collection is existing or not
            if (isSiteCollection)
            {
                try
                {
                    var properties = tenant.GetSitePropertiesByUrl(siteFullUrl, false);
                    tenant.Context.Load(properties);
                    tenant.Context.ExecuteQueryRetry();
                    ret = properties.Status.Equals(status, StringComparison.OrdinalIgnoreCase);
                }
                catch (ServerException ex)
                {
                    if (IsUnableToAccessSiteException(ex))
                    {
                        try
                        {
                            //Let's retry to see if this site collection was recycled
                            var deletedProperties = tenant.GetDeletedSitePropertiesByUrl(siteFullUrl);
                            tenant.Context.Load(deletedProperties);
                            tenant.Context.ExecuteQueryRetry();
                            ret = deletedProperties.Status.Equals(status, StringComparison.OrdinalIgnoreCase);
                        }
                        catch
                        {
                            // eat exception
                        }
                    }
                }
            }
            //Judge whether this sub web site is existing or not
            else
            {
                var subsiteUrl = string.Format(CultureInfo.CurrentCulture,
                            "{0}{1}{2}", siteDomainUrl, managedPath, siteRelativePath.Split('/')[0]);
                var subsiteRelativeUrl = siteRelativePath.Substring(siteRelativePath.IndexOf('/') + 1);
                var site = tenant.GetSiteByUrl(subsiteUrl);
                var subweb = site.OpenWeb(subsiteRelativeUrl);
                tenant.Context.Load(subweb, w => w.Title);
                tenant.Context.ExecuteQueryRetry();
                ret = true;
            }
            return ret;
        }

        /// <summary>
        /// Checks if a site collection is Active
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">URL to the site collection</param>
        /// <returns>True if active, false if not</returns>
        public static bool IsSiteActive(this Tenant tenant, string siteFullUrl)
        {
            try
            {
                return tenant.CheckIfSiteExists(siteFullUrl, "Active");
            }
            catch (Exception ex)
            {
                if (IsCannotGetSiteException(ex))
                {
                    return false;
                }

                Log.Error(CoreResources.TenantExtensions_UnknownExceptionAccessingSite, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Checks if a site collection exists, relies on tenant admin API. Sites that are recycled also return as existing sites
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">URL to the site collection</param>
        /// <returns>True if existing, false if not</returns>
        public static bool SiteExists(this Tenant tenant, string siteFullUrl)
        {
            try
            {
                //Get the site name
                var properties = tenant.GetSitePropertiesByUrl(siteFullUrl, false);
                tenant.Context.Load(properties);
                tenant.Context.ExecuteQueryRetry();

                // Will cause an exception if site URL is not there. Not optimal, but the way it works.
                return true;
            }
            catch (Exception ex)
            {
                if (IsCannotGetSiteException(ex) || IsUnableToAccessSiteException(ex))
                {
                    if (IsUnableToAccessSiteException(ex))
                    {
                        //Let's retry to see if this site collection was recycled
                        try
                        {
                            var deletedProperties = tenant.GetDeletedSitePropertiesByUrl(siteFullUrl);
                            tenant.Context.Load(deletedProperties);
                            tenant.Context.ExecuteQueryRetry();
                            return deletedProperties.Status.Equals("Recycled", StringComparison.OrdinalIgnoreCase);
                        }
                        catch
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// Checks if a sub site exists
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">URL to the sub site</param>
        /// <returns>True if existing, false if not</returns>
        public static bool SubSiteExists(this Tenant tenant, string siteFullUrl)
        {
            try
            {
                return tenant.CheckIfSiteExists(siteFullUrl, "Active");
            }
            catch (Exception ex)
            {
                if (IsCannotGetSiteException(ex) || IsUnableToAccessSiteException(ex))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        #endregion

        #region Site collection deletion
        /// <summary>
        /// Deletes a site collection
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">Url of the site collection to delete</param>
        /// <param name="useRecycleBin">Leave the deleted site collection in the site collection recycle bin</param>
        /// <param name="timeoutFunction">An optional function that will be called while waiting for the site to be created. Return true to cancel the wait loop.</param>
        /// <returns>True if deleted</returns>
        public static bool DeleteSiteCollection(this Tenant tenant, string siteFullUrl, bool useRecycleBin, Func<TenantOperationMessage, bool> timeoutFunction = null)
        {
            var succeeded = false;
            bool ret = false;

            try
            {
                SpoOperation op = tenant.RemoveSite(siteFullUrl);
                tenant.Context.Load(tenant);
                tenant.Context.Load(op, i => i.IsComplete, i => i.PollingInterval);
                tenant.Context.ExecuteQueryRetry();

                //check if site creation operation is complete
                succeeded = WaitForIsComplete(tenant, op, timeoutFunction, TenantOperationMessage.DeletingSiteCollection);
            }
            catch (ServerException ex)
            {
                if (!useRecycleBin && IsCannotRemoveSiteException(ex))
                {
                    //eat exception as the site might be in the recycle bin and we allowed deletion from recycle bin 
                }
                else
                {
                    throw;
                }
            }

            if (useRecycleBin)
            {
                return true;
            }

            if (succeeded)
            {
                // To delete Site collection completely, (may take a longer time)
                SpoOperation op2 = tenant.RemoveDeletedSite(siteFullUrl);
                tenant.Context.Load(op2, i => i.IsComplete, i => i.PollingInterval);
                tenant.Context.ExecuteQueryRetry();

                succeeded = WaitForIsComplete(tenant, op2, timeoutFunction,
                    TenantOperationMessage.RemovingDeletedSiteCollectionFromRecycleBin);
                ret = succeeded;
            }
            return ret;
        }

        /// <summary>
        /// Deletes a site collection from the site collection recycle bin
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">URL of the site collection to delete</param>
        /// <param name="wait">If true, processing will halt until the site collection has been deleted from the recycle bin</param>
        /// <param name="timeoutFunction">An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.</param>
        public static bool DeleteSiteCollectionFromRecycleBin(this Tenant tenant, string siteFullUrl, bool wait = true, Func<TenantOperationMessage, bool> timeoutFunction = null)
        {
            var ret = true;
            var op = tenant.RemoveDeletedSite(siteFullUrl);
            tenant.Context.Load(op, i => i.IsComplete, i => i.PollingInterval);
            tenant.Context.ExecuteQueryRetry();
            if (timeoutFunction != null)
            {
                wait = true;
            }
            if (wait)
            {
                var succeeded = WaitForIsComplete(tenant, op, timeoutFunction,
                    TenantOperationMessage.RemovingDeletedSiteCollectionFromRecycleBin);
                ret = succeeded;
            }
            return ret;
        }
        #endregion

        #region Site collection properties
        /// <summary>
        /// Gets the ID of site collection with specified URL
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">A URL that specifies a site collection to get ID.</param>
        /// <returns>The Guid of a site collection</returns>
        public static Guid GetSiteGuidByUrl(this Tenant tenant, string siteFullUrl)
        {
            if (string.IsNullOrEmpty(siteFullUrl))
                throw new ArgumentNullException("siteFullUrl");

            return tenant.GetSiteGuidByUrl(new Uri(siteFullUrl));
        }

        /// <summary>
        /// Gets the ID of site collection with specified URL
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">A URL that specifies a site collection to get ID.</param>
        /// <returns>The Guid of a site collection or an Guid.Empty if the Site does not exist</returns>
        public static Guid GetSiteGuidByUrl(this Tenant tenant, Uri siteFullUrl)
        {
            Site site = null;
            site = tenant.GetSiteByUrl(siteFullUrl.OriginalString);
            tenant.Context.Load(site);
            tenant.Context.ExecuteQueryRetry();
            var siteGuid = site.Id;

            return siteGuid;
        }

        /// <summary>
        /// Returns available webtemplates/site definitions
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="lcid">Locale identifier (LCID) for the language</param>
        /// <param name="compatibilityLevel">14 for SharePoint 2010, 15 for SharePoint 2013/SharePoint Online</param>
        /// <returns>Returns collection of SPTenantWebTemplate</returns>
        public static SPOTenantWebTemplateCollection GetWebTemplates(this Tenant tenant, uint lcid, int compatibilityLevel)
        {

            var templates = tenant.GetSPOTenantWebTemplates(lcid, compatibilityLevel);

            tenant.Context.Load(templates);

            tenant.Context.ExecuteQueryRetry();

            return templates;
        }

        /// <summary>
        /// Sets tenant site Properties
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">full URL of site</param>
        /// <param name="title">site title</param>
        /// <param name="allowSelfServiceUpgrade">Boolean value to allow serlf service upgrade</param>
        /// <param name="sharingCapability">SharingCapabilities enumeration value (i.e. Disabled/ExternalUserSharingOnly/ExternalUserAndGuestSharing/ExistingExternalUserSharingOnly)</param>
        /// <param name="storageMaximumLevel">A limit on all disk space used by the site collection</param>
        /// <param name="storageWarningLevel">A storage warning level for when administrators of the site collection receive advance notice before available storage is expended.</param>
        /// <param name="userCodeMaximumLevel">A value that represents the maximum allowed resource usage for the site/</param>
        /// <param name="userCodeWarningLevel">A value that determines the level of resource usage at which a warning e-mail message is sent</param>
        /// <param name="noScriptSite">Boolean value which allows to customize the site using scripts</param>
        /// <param name="wait">Id true this function only returns when the tenant properties are set, if false it will return immediately</param>
        /// <param name="timeoutFunction">An optional function that will be called while waiting for the tenant properties to be set. If set will override the wait variable. Return true to cancel the wait loop.</param>
        public static void SetSiteProperties(this Tenant tenant, string siteFullUrl,
            string title = null,
            bool? allowSelfServiceUpgrade = null,
            SharingCapabilities? sharingCapability = null,
            long? storageMaximumLevel = null,
            long? storageWarningLevel = null,
            double? userCodeMaximumLevel = null,
            double? userCodeWarningLevel = null,
            bool? noScriptSite = null,
            bool wait = true, Func<TenantOperationMessage, bool> timeoutFunction = null
            )
        {
            var siteProps = tenant.GetSitePropertiesByUrl(siteFullUrl, true);
            tenant.Context.Load(siteProps);
            tenant.Context.ExecuteQueryRetry();
            if (siteProps != null)
            {
                if (allowSelfServiceUpgrade != null)
                    siteProps.AllowSelfServiceUpgrade = allowSelfServiceUpgrade.Value;
                if (sharingCapability != null)
                    siteProps.SharingCapability = sharingCapability.Value;
                if (storageMaximumLevel != null)
                    siteProps.StorageMaximumLevel = storageMaximumLevel.Value;
                if (storageWarningLevel != null)
                    siteProps.StorageWarningLevel = storageWarningLevel.Value;
                if (userCodeMaximumLevel != null)
                    siteProps.UserCodeMaximumLevel = userCodeMaximumLevel.Value;
                if (userCodeWarningLevel != null)
                    siteProps.UserCodeWarningLevel = userCodeWarningLevel.Value;
                if (title != null)
                    siteProps.Title = title;
                if (noScriptSite != null)
                    siteProps.DenyAddAndCustomizePages = (noScriptSite == true ? DenyAddAndCustomizePagesStatus.Enabled : DenyAddAndCustomizePagesStatus.Disabled);

                var op = siteProps.Update();
                tenant.Context.Load(op, i => i.IsComplete, i => i.PollingInterval);
                tenant.Context.ExecuteQueryRetry();
                if (timeoutFunction != null)
                {
                    wait = true;
                }
                if (wait)
                {
                    WaitForIsComplete(tenant, op, timeoutFunction, TenantOperationMessage.SettingSiteProperties);
                }
            }
        }

        /// <summary>
        /// Sets a site to Unlock access or NoAccess. This operation may occur immediately, but the site lock may take a short while before it goes into effect.
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site (i.e. https://[tenant]-admin.sharepoint.com)</param>
        /// <param name="siteFullUrl">The target site to change the lock state.</param>
        /// <param name="lockState">The target state the site should be changed to.</param>
        /// <param name="wait">If true, processing will halt until the site collection lock state has been implemented</param>      
        /// <param name="timeoutFunction">An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.</param>
        public static void SetSiteLockState(this Tenant tenant, string siteFullUrl, SiteLockState lockState, bool wait = false, Func<TenantOperationMessage, bool> timeoutFunction = null)
        {
            var siteProps = tenant.GetSitePropertiesByUrl(siteFullUrl, false);
            tenant.Context.Load(siteProps);
            tenant.Context.ExecuteQueryRetry();

            Log.Info(CoreResources.TenantExtensions_SetLockState, siteProps.LockState, lockState);

            if (siteProps.LockState != lockState.ToString())
            {
                siteProps.LockState = lockState.ToString();
                SpoOperation op = siteProps.Update();
                tenant.Context.Load(op, i => i.IsComplete, i => i.PollingInterval);
                tenant.Context.ExecuteQueryRetry();
                if (timeoutFunction != null)
                {
                    wait = true;
                }
                if (wait)
                {
                    WaitForIsComplete(tenant, op, timeoutFunction, TenantOperationMessage.SettingSiteLockState);
                }

            }
        }
        #endregion

        #region Site collection administrators
        /// <summary>
        /// Add a site collection administrator to a site collection
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="adminLogins">Array of admins loginnames to add</param>
        /// <param name="siteUrl">Url of the site to operate on</param>
        /// <param name="addToOwnersGroup">Optionally the added admins can also be added to the Site owners group</param>
        public static void AddAdministrators(this Tenant tenant, IEnumerable<UserEntity> adminLogins, Uri siteUrl, bool addToOwnersGroup = false)
        {
            if (adminLogins == null)
                throw new ArgumentNullException("adminLogins");

            if (siteUrl == null)
                throw new ArgumentNullException("siteUrl");

            foreach (UserEntity admin in adminLogins)
            {
                var siteUrlString = siteUrl.ToString();
                tenant.SetSiteAdmin(siteUrlString, admin.LoginName, true);
                tenant.Context.ExecuteQueryRetry();
                if (addToOwnersGroup)
                {
                    // Create a separate context to the web
                    using (var clientContext = tenant.Context.Clone(siteUrl))
                    {
                        var spAdmin = clientContext.Web.EnsureUser(admin.LoginName);
                        clientContext.Web.AssociatedOwnerGroup.Users.AddUser(spAdmin);
                        clientContext.Web.AssociatedOwnerGroup.Update();
                        clientContext.ExecuteQueryRetry();
                    }
                }
            }
        }
        #endregion

        #region Site enumeration
        /// <summary>
        /// Returns all site collections in the current Tenant based on a startIndex. IncludeDetail adds additional properties to the SPSite object. 
        /// </summary>
        /// <param name="tenant">Tenant object to operate against</param>
        /// <param name="startIndex">Not relevant anymore</param>
        /// <param name="endIndex">Not relevant anymore</param>
        /// <param name="includeDetail">Option to return a limited set of data</param>
        /// <param name="includeOD4BSites">Also return the OD4B sites</param>
        /// <returns>An IList of SiteEntity objects</returns>
        public static IList<SiteEntity> GetSiteCollections(this Tenant tenant, int startIndex = 0, int endIndex = 500000, bool includeDetail = true, bool includeOD4BSites = false)
        {
            var sites = new List<SiteEntity>();
            SPOSitePropertiesEnumerable props = null;

            while (props == null || props.NextStartIndexFromSharePoint != null)
            {

                // approach to be used as of Feb 2017
                SPOSitePropertiesEnumerableFilter filter = new SPOSitePropertiesEnumerableFilter()
                {
                    IncludePersonalSite = includeOD4BSites ? PersonalSiteFilter.Include : PersonalSiteFilter.UseServerDefault,
                    StartIndex = props == null ? null : props.NextStartIndexFromSharePoint,
                    IncludeDetail = includeDetail
                };
                props = tenant.GetSitePropertiesFromSharePointByFilters(filter);

                // Previous approach, being replaced by GetSitePropertiesFromSharePointByFilters which also allows to fetch OD4B sites
                //props = tenant.GetSitePropertiesFromSharePoint(props == null ? null : props.NextStartIndexFromSharePoint, includeDetail);
                tenant.Context.Load(props);
                tenant.Context.ExecuteQueryRetry();

                foreach (var prop in props)
                {
                    var siteEntity = new SiteEntity();
                    siteEntity.Lcid = prop.Lcid;
                    siteEntity.SiteOwnerLogin = prop.Owner;
                    siteEntity.StorageMaximumLevel = prop.StorageMaximumLevel;
                    siteEntity.StorageWarningLevel = prop.StorageWarningLevel;
                    siteEntity.Template = prop.Template;
                    siteEntity.TimeZoneId = prop.TimeZoneId;
                    siteEntity.Title = prop.Title;
                    siteEntity.Url = prop.Url;
                    siteEntity.UserCodeMaximumLevel = prop.UserCodeMaximumLevel;
                    siteEntity.UserCodeWarningLevel = prop.UserCodeWarningLevel;
                    siteEntity.CurrentResourceUsage = prop.CurrentResourceUsage;
                    siteEntity.LastContentModifiedDate = prop.LastContentModifiedDate;
                    siteEntity.StorageUsage = prop.StorageUsage;
                    siteEntity.WebsCount = prop.WebsCount;
                    SiteLockState lockState;
                    if (Enum.TryParse(prop.LockState, out lockState))
                    {
                        siteEntity.LockState = lockState;
                    }
                    sites.Add(siteEntity);
                }
            }

            return sites;
        }

#if !NETSTANDARD2_0
        /// <summary>
        /// Get OneDrive site collections by iterating through all user profiles.
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site </param>
        /// <returns>List of <see cref="SiteEntity"/> objects containing site collection info.</returns>
        public static IList<SiteEntity> GetOneDriveSiteCollections(this Tenant tenant)
        {
            var sites = new List<SiteEntity>();
            var svcClient = GetUserProfileServiceClient(tenant);

            // get all user profiles
            var userProfileResult = svcClient.GetUserProfileByIndex(-1);

            while (int.Parse(userProfileResult.NextValue) != -1)
            {
                var personalSpaceProperty = userProfileResult.UserProfile.FirstOrDefault(p => p.Name == "PersonalSpace");

                if (personalSpaceProperty != null && personalSpaceProperty.Values.Any())
                {
                    var usernameProperty = userProfileResult.UserProfile.FirstOrDefault(p => p.Name == "UserName");
                    var nameProperty = userProfileResult.UserProfile.FirstOrDefault(p => p.Name == "PreferredName");
                    var url = personalSpaceProperty.Values[0].Value as string;
                    var name = nameProperty.Values[0].Value as string;
                    var siteEntity = new SiteEntity
                    {
                        Url = url,
                        Title = name,
                        SiteOwnerLogin = usernameProperty.Values[0].Value as string
                    };
                    sites.Add(siteEntity);
                }

                userProfileResult = svcClient.GetUserProfileByIndex(int.Parse(userProfileResult.NextValue));
            }

            return sites;
        }
#endif

#if !NETSTANDARD2_0
        /// <summary>
        /// Gets the UserProfileService proxy to enable calls to the UPA web service.
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site </param>
        /// <returns>UserProfileService web service client</returns>
        public static UserProfileService GetUserProfileServiceClient(this Tenant tenant)
        {
            var client = new UserProfileService();

            client.Url = tenant.Context.Url + "/_vti_bin/UserProfileService.asmx";
            client.UseDefaultCredentials = false;
            client.Credentials = tenant.Context.Credentials;

            if (tenant.Context.Credentials is SharePointOnlineCredentials)
            {
                var creds = (SharePointOnlineCredentials)tenant.Context.Credentials;
                var authCookie = creds.GetAuthenticationCookie(new Uri(tenant.Context.Url));
                var cookieContainer = new CookieContainer();

                cookieContainer.SetCookies(new Uri(tenant.Context.Url), authCookie);
                client.CookieContainer = cookieContainer;
            }
            return client;
        }
#endif

        #endregion

        #region ClientSide Package Deployment

        /// <summary>
        /// Gets the Uri for the tenant's app catalog site (if that one has already been created)
        /// </summary>
        /// <param name="tenant">Tenant to operate against</param>
        /// <returns>The Uri holding the app catalog site URL</returns>
        public static Uri GetAppCatalog(this Tenant tenant)
        {
            // Assume there's only one appcatalog site
            var results = ((tenant.Context) as ClientContext).Web.SiteSearch("contentclass:STS_Site AND SiteTemplate:APPCATALOG");
            foreach (var site in results)
            {
                return new Uri(site.Url);
            }

            return null;
        }
        #endregion

        #region Private helper methods
        private static bool WaitForIsComplete(Tenant tenant, SpoOperation op, Func<TenantOperationMessage, bool> timeoutFunction = null, TenantOperationMessage operationMessage = TenantOperationMessage.None)
        {
            bool succeeded = true;
            while (!op.IsComplete)
            {
                if (timeoutFunction != null && timeoutFunction(operationMessage))
                {
                    succeeded = false;
                    break;
                }
                Thread.Sleep(op.PollingInterval);

                op.RefreshLoad();
                if (!op.IsComplete)
                {
                    try
                    {
                        tenant.Context.ExecuteQueryRetry();
                    }
                    catch (WebException webEx)
                    {
                        // Context connection gets closed after action completed.
                        // Calling ExecuteQuery again returns an error which can be ignored
                        Log.Warning(CoreResources.TenantExtensions_ClosedContextWarning, webEx.Message);
                    }
                }
            }
            return succeeded;
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
                if (
                     (((ServerException)ex).ServerErrorCode == -2147024809 && ((ServerException)ex).ServerErrorTypeName.Equals("System.ArgumentException", StringComparison.InvariantCultureIgnoreCase)) ||
                     (((ServerException)ex).ServerErrorCode == -1 && ((ServerException)ex).ServerErrorTypeName.Equals("Microsoft.Online.SharePoint.Common.SpoNoSiteException", StringComparison.InvariantCultureIgnoreCase))
                    )
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

        private static bool IsCannotRemoveSiteException(Exception ex)
        {
            if (ex is ServerException)
            {
                if (((ServerException)ex).ServerErrorCode == -1
                    && (
                        ((ServerException)ex).ServerErrorTypeName.Equals("Microsoft.Online.SharePoint.Common.SpoException", StringComparison.InvariantCultureIgnoreCase) ||
                        ((ServerException)ex).ServerErrorTypeName.Equals("Microsoft.Online.SharePoint.Common.SpoNoSiteException", StringComparison.InvariantCultureIgnoreCase))
                    )
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

        #region Site Classification configuration

        /// <summary>
        /// Enables Site Classifications for the target tenant 
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="siteClassificationsSettings">The site classifications settings to apply./param>
        public static void EnableSiteClassifications(this Tenant tenant, string accessToken, SiteClassificationsSettings siteClassificationsSettings)
        {
            SiteClassificationsUtility.EnableSiteClassifications(accessToken, siteClassificationsSettings);
        }

        /// <summary>
        /// Enables Site Classifications for the target tenant 
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="classificationsList">The list of classification values</param>
        /// <param name="defaultClassification">The default classification</param>
        /// <param name="usageGuidelinesUrl">The URL of a guidance page</param>
        public static void EnableSiteClassifications(this Tenant tenant, string accessToken, IEnumerable<string> classificationsList, string defaultClassification = "", string usageGuidelinesUrl = "")
        {
            SiteClassificationsUtility.EnableSiteClassifications(accessToken, classificationsList, defaultClassification, usageGuidelinesUrl);
        }

        /// <summary>
        /// Enables Site Classifications for the target tenant 
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <returns>The list of Site Classifications values</returns>
        public static SiteClassificationsSettings GetSiteClassificationsSettings(this Tenant tenant, string accessToken)
        {
            return SiteClassificationsUtility.GetSiteClassificationsSettings(accessToken);
        }

        /// <summary>
        /// Updates Site Classifications settings for the target tenant
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="siteClassificationsSettings">The site classifications settings to update.</param>
        public static void UpdateSiteClassificationsSettings(this Tenant tenant, string accessToken, SiteClassificationsSettings siteClassificationsSettings)
        {
            SiteClassificationsUtility.UpdateSiteClassificationsSettings(accessToken, siteClassificationsSettings);
        }

        /// <summary>
        /// Updates Site Classifications settings for the target tenant
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="classificationsList">The list of classification values</param>
        /// <param name="defaultClassification">The default classification</param>
        /// <param name="usageGuidelinesUrl">The URL of a guidance page</param>
        public static void UpdateSiteClassificationsSettings(this Tenant tenant, string accessToken, IEnumerable<string> classificationsList, string defaultClassification = "", string usageGuidelinesUrl = "")
        {
            SiteClassificationsUtility.UpdateSiteClassificationsSettings(accessToken, classificationsList, defaultClassification, usageGuidelinesUrl);
        }

        /// <summary>
        /// Disables Site Classifications settings for the target tenant
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        public static void DisableSiteClassifications(this Tenant tenant, string accessToken)
        {
            SiteClassificationsUtility.DisableSiteClassifications(accessToken);
        }

        #endregion

        #region Site groupify
        /// <summary>
        /// Connect an Office 365 group to an existing SharePoint site collection
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="siteUrl">Url to the site collection that needs to get connected to an Office 365 group</param>
        /// <param name="siteCollectionGroupifyInformation">Information that configures the "groupify" process</param>
        public static void GroupifySite(this Tenant tenant, string siteUrl, TeamSiteCollectionGroupifyInformation siteCollectionGroupifyInformation)
        {
            if (string.IsNullOrEmpty(siteUrl))
            {
                throw new ArgumentException("Missing value for siteUrl", "siteUrl");
            }

            if (siteCollectionGroupifyInformation == null)
            {
                throw new ArgumentException("Missing value for siteCollectionGroupifyInformation", "sitecollectionGroupifyInformation");
            }

            if (!string.IsNullOrEmpty(siteCollectionGroupifyInformation.Alias) && siteCollectionGroupifyInformation.Alias.Contains(" "))
            {
                throw new ArgumentException("Alias cannot contain spaces", "Alias");
            }

            if (string.IsNullOrEmpty(siteCollectionGroupifyInformation.DisplayName))
            {
                throw new ArgumentException("DisplayName is required", "DisplayName");
            }

            GroupCreationParams optionalParams = new GroupCreationParams(tenant.Context);
            if (!String.IsNullOrEmpty(siteCollectionGroupifyInformation.Description))
            {
                optionalParams.Description = siteCollectionGroupifyInformation.Description;
            }
            if (!String.IsNullOrEmpty(siteCollectionGroupifyInformation.Classification))
            {
                optionalParams.Classification = siteCollectionGroupifyInformation.Classification;
            }
            if (siteCollectionGroupifyInformation.KeepOldHomePage)
            {
                optionalParams.CreationOptions = new string[] { "SharePointKeepOldHomepage" };
            }

            tenant.CreateGroupForSite(siteUrl, siteCollectionGroupifyInformation.DisplayName, siteCollectionGroupifyInformation.Alias, siteCollectionGroupifyInformation.IsPublic, optionalParams);
            tenant.Context.ExecuteQueryRetry();
        }
        #endregion

#else
        #region Site collection creation
        /// <summary>
        /// Adds a SiteEntity by launching site collection creation and waits for the creation to finish
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="properties">Describes the site collection to be created</param>
        public static void CreateSiteCollection(this Tenant tenant, SiteEntity properties)
        {
            SiteCreationProperties newsite = new SiteCreationProperties();
            newsite.Url = properties.Url;
            newsite.Owner = properties.SiteOwnerLogin;
            newsite.Template = properties.Template;
            newsite.Title = properties.Title;
            newsite.StorageMaximumLevel = properties.StorageMaximumLevel;
            newsite.StorageWarningLevel = properties.StorageWarningLevel;
            newsite.TimeZoneId = properties.TimeZoneId;
            newsite.UserCodeMaximumLevel = properties.UserCodeMaximumLevel;
            newsite.UserCodeWarningLevel = properties.UserCodeWarningLevel;
            newsite.Lcid = properties.Lcid;

            try
            {
                tenant.CreateSite(newsite);
                tenant.Context.ExecuteQueryRetry();
            }
            catch (Exception ex)
            {
                // Eat the siteSubscription exception to make the same code work for MT as on-prem April 2014 CU+
                if (ex.Message.IndexOf("Parameter name: siteSubscription") == -1)
                {
                    throw;
                }
            }
        }
#endregion

#region Site collection deletion
        /// <summary>
        /// Deletes a site collection
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">Url of the site collection to delete</param>
        public static void DeleteSiteCollection(this Tenant tenant, string siteFullUrl)
        {
            tenant.RemoveSite(siteFullUrl);
            tenant.Context.ExecuteQueryRetry();
        }
#endregion
#endif
    }
}
