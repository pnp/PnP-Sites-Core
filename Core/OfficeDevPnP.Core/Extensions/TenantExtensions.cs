using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
#if !NETSTANDARD2_0
#endif
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
#if !NETSTANDARD2_0
using OfficeDevPnP.Core.UPAWebService;
#endif
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Graph;
using OfficeDevPnP.Core.Framework.Graph.Model;
#if !ONPREMISES
using OfficeDevPnP.Core.Sites;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Utilities;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using System.Threading.Tasks;
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
        #region Provisioning

        /// <summary>
        /// Applies a template to a tenant
        /// </summary>
        /// <param name="tenant"></param>
        /// <param name="tenantTemplate"></param>
        /// <param name="sequenceId"></param>
        /// <param name="configuration"></param>
        public static void ApplyTenantTemplate(this Tenant tenant, ProvisioningHierarchy tenantTemplate, string sequenceId, ApplyConfiguration configuration = null)
        {
            SiteToTemplateConversion engine = new SiteToTemplateConversion();
            engine.ApplyTenantTemplate(tenant, tenantTemplate, sequenceId, configuration);
        }

        /// <summary>
        /// Extracts a template from a tenant
        /// </summary>
        /// <param name="tenant"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public static ProvisioningHierarchy GetTenantTemplate(this Tenant tenant, ExtractConfiguration configuration)
        {
            return new SiteToTemplateConversion().GetTenantTemplate(tenant, configuration);
        }

        /// <summary>
        /// Returns the urls of sites connected to the hubsite specified
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="hubSiteUrl">The fully qualified url of the hubsite</param>
        /// <returns></returns>
        public static List<string> GetHubSiteChildUrls(this Tenant tenant, string hubSiteUrl)
        {
            var properties = tenant.GetHubSitePropertiesByUrl(hubSiteUrl);
            tenant.Context.Load(properties);
            tenant.Context.ExecuteQueryRetry();
            return GetHubSiteChildUrls(tenant, properties.ID);
        }

        /// <summary>
        /// Returns the urls of sites connected to the hubsite specified
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="hubsiteId">The id of the hubsite</param>
        /// <returns></returns>
        public static List<string> GetHubSiteChildUrls(this Tenant tenant, Guid hubsiteId)
        {
            List<string> urls = new List<string>();
            using (var tenantContext = tenant.Context.Clone((tenant.Context as ClientContext).Web.GetTenantAdministrationUrl()))
            {
                var siteList = tenantContext.Web.Lists.GetByTitle("DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS");
                var query = new CamlQuery()
                {
                    ViewXml = $"<View><Query><Where><And><Eq><FieldRef Name='HubSiteId' /><Value Type='Guid'>{hubsiteId}</Value></Eq><And><Neq><FieldRef Name='SiteId' /><Value Type='Guid'>{hubsiteId}</Value></Neq><IsNull><FieldRef Name='TimeDeleted'/></IsNull></And></And></Where></Query><ViewFields><FieldRef Name='SiteUrl'/></ViewFields></View><RowLimit Paging='TRUE'>100</RowLimit>"
                };

                do
                {
                    var items = siteList.GetItems(query);
                    tenantContext.Load(items);
                    tenantContext.ExecuteQueryRetry();
                    foreach (var item in items)
                    {
                        urls.Add(item["SiteUrl"].ToString());
                    }
                    query.ListItemCollectionPosition = items.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null);
            }
            return urls;
        }
        #endregion

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

        /// <summary>
        /// Creates a new App Catalog and registers the app catalog site as the tenant App Catalog.
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="url">The Full Site Url, e.g. https://yourtenant.sharepoint.com/sites/appcatalog</param>
        /// <param name="ownerLogin">The username of the owner of the appcatalog, e.g. user@domain.com</param>
        /// <param name="timeZoneId">TimeZoneId for the appcatalog site. "(UTC+01:00) Brussels, Copenhagen, Madrid, Paris" = 3"</param>
        /// <param name="force">If true, and an appcatalog is already registered and present, the new appcatalog will be created. If the same URL is provided and the site is present the current one will be deleted and a new one will be created.</param>
        /// <returns></returns>
        public static async Task EnsureAppCatalogAsync(this Tenant tenant, string url, string ownerLogin, int timeZoneId, bool force = false)
        {

            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentException("App Catalog Site Url is required", nameof(url));
            }

            if (string.IsNullOrEmpty(ownerLogin))
            {
                throw new ArgumentException("Owner is required", nameof(ownerLogin));
            }

            // Check if there is already an app catalog
            var settings = TenantSettings.GetCurrent(tenant.Context);
            var appCatalogUrl = await settings.EnsurePropertyAsync(s => s.CorporateCatalogUrl);
            if (!string.IsNullOrEmpty(appCatalogUrl))
            {
                // check if the site exists
                var siteExistence = tenant.SiteExistsAnywhere(appCatalogUrl);
                if (siteExistence == SiteExistence.No)
                {
                    CreateAppCatalogInternal(tenant, url, ownerLogin, timeZoneId, force);
                }
                else if (force)
                {
                    DeleteSiteCollection(tenant, appCatalogUrl, false);
                    CreateAppCatalogInternal(tenant, url, ownerLogin, timeZoneId, force);
                } else
                {
                    throw new Exception($"An App Catalog already exists at {appCatalogUrl} and force is not specified.");
                }
            }
            else
            {
                CreateAppCatalogInternal(tenant, url, ownerLogin, timeZoneId, true);
            }
        }

        private static void CreateAppCatalogInternal(Tenant tenant, string url, string ownerLogin, int timeZoneId, bool removeFromRecycleBin)
        {
            var siteEntity = new SiteEntity
            {
                Template = "APPCATALOG#0",
                SiteOwnerLogin = ownerLogin,
                TimeZoneId = timeZoneId,
                Url = url
            };
            CreateSiteCollection(tenant, siteEntity, removeFromRecycleBin, true);
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
#endif

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

#if !SP2013
#if ONPREMISES
        /// <summary>
        /// Sets tenant site Properties
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">full URL of site</param>
        /// <param name="title">site title</param>
        /// <param name="allowSelfServiceUpgrade">Boolean value to allow serlf service upgrade</param>
        /// <param name="noScriptSite">Boolean value which allows to customize the site using scripts</param>
        public static void SetSiteProperties(this Tenant tenant, string siteFullUrl,
            string title = null,
            bool? allowSelfServiceUpgrade = null,
            bool? noScriptSite = null
        )
#else
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
        /// <param name="commentsOnSitePagesDisabled">Boolean value which Enables/Disables comments on the Site Pages</param>
        /// <param name="socialBarOnSitePagesDisabled">Boolean value which Enables/Disables likes and view count on the Site Pages</param>
        /// <param name="defaultSharingLinkType">Specifies the default link type for the site collection</param>
        /// <param name="wait">Id true this function only returns when the tenant properties are set, if false it will return immediately</param>
        /// <param name="timeoutFunction">An optional function that will be called while waiting for the tenant properties to be set. If set will override the wait variable. Return true to cancel the wait loop.</param>
        /// <param name="defaultLinkPermission">Specifies the default link permission for the site collection</param>
        public static void SetSiteProperties(this Tenant tenant, string siteFullUrl,
            string title = null,
            bool? allowSelfServiceUpgrade = null,
            SharingCapabilities? sharingCapability = null,
            long? storageMaximumLevel = null,
            long? storageWarningLevel = null,
            double? userCodeMaximumLevel = null,
            double? userCodeWarningLevel = null,
            bool? noScriptSite = null,
            bool? commentsOnSitePagesDisabled = null,
            bool? socialBarOnSitePagesDisabled = null,
            Microsoft.Online.SharePoint.TenantManagement.SharingPermissionType? defaultLinkPermission = null,
            Microsoft.Online.SharePoint.TenantManagement.SharingLinkType? defaultSharingLinkType = null,
            bool wait = true, Func<TenantOperationMessage, bool> timeoutFunction = null
            )
#endif
        {
            var siteProps = tenant.GetSitePropertiesByUrl(siteFullUrl, true);
            tenant.Context.Load(siteProps);
            tenant.Context.ExecuteQueryRetry();
            if (siteProps != null)
            {
                if (allowSelfServiceUpgrade != null)
                    siteProps.AllowSelfServiceUpgrade = allowSelfServiceUpgrade.Value;
#if !ONPREMISES
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
                if (defaultLinkPermission != null)
                    siteProps.DefaultLinkPermission = defaultLinkPermission.Value;
                if (defaultSharingLinkType != null)
                    siteProps.DefaultSharingLinkType = defaultSharingLinkType.Value;
#endif
                if (title != null)
                    siteProps.Title = title;
                if (noScriptSite != null)
                    siteProps.DenyAddAndCustomizePages = (noScriptSite == true ? DenyAddAndCustomizePagesStatus.Enabled : DenyAddAndCustomizePagesStatus.Disabled);
#if !ONPREMISES
                if (commentsOnSitePagesDisabled != null)
                    siteProps.CommentsOnSitePagesDisabled = commentsOnSitePagesDisabled.Value;
                if (socialBarOnSitePagesDisabled != null)
                    siteProps.SocialBarOnSitePagesDisabled = socialBarOnSitePagesDisabled.Value;
#endif

                var op = siteProps.Update();
                tenant.Context.Load(op, i => i.IsComplete, i => i.PollingInterval);
                tenant.Context.ExecuteQueryRetry();
#if !ONPREMISES
                if (timeoutFunction != null)
                {
                    wait = true;
                }
                if (wait)
                {
                    WaitForIsComplete(tenant, op, timeoutFunction, TenantOperationMessage.SettingSiteProperties);
                }
#endif
            }
        }
#endif

#if !ONPREMISES
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
#endif
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

#if ONPREMISES
#if !SP2013
        /// <summary>
        /// Returns all site collections in the current Tenant based on a startIndex. IncludeDetail adds additional properties to the SPSite object. 
        /// </summary>
        /// <param name="tenant">Tenant object to operate against</param>
        /// <param name="startIndex">Start getting site collections from this index. Defaults to 0</param>
        /// <param name="includeDetail">Option to return a limited set of data</param>
        /// <returns>An IList of SiteEntity objects</returns>
        public static IList<SiteEntity> GetSiteCollections(this Tenant tenant, int startIndex = 0, bool includeDetail = true)
        {
            var sites = new List<SiteEntity>();

            SPOSitePropertiesEnumerable props = tenant.GetSiteProperties(startIndex, includeDetail);
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

            return sites;
        }
#endif
#else
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
#endif

        #endregion

#if !ONPREMISES
        

        #region Site Classification configuration

        /// <summary>
        /// Enables Site Classifications for the target tenant 
        /// </summary>
        /// <param name="tenant">The target tenant</param>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="siteClassificationsSettings">The site classifications settings to apply.</param>
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

            var creationOptionsValues = new List<string>();
            if (siteCollectionGroupifyInformation.KeepOldHomePage)
            {
                creationOptionsValues.Add("SharePointKeepOldHomepage");
            }
            creationOptionsValues.Add($"HubSiteId:{siteCollectionGroupifyInformation.HubSiteId}");
            optionalParams.CreationOptions = creationOptionsValues.ToArray();

            if (siteCollectionGroupifyInformation.Owners != null && siteCollectionGroupifyInformation.Owners.Length > 0)
            {
                optionalParams.Owners = siteCollectionGroupifyInformation.Owners;
            }

            tenant.CreateGroupForSite(siteUrl, siteCollectionGroupifyInformation.DisplayName, siteCollectionGroupifyInformation.Alias, siteCollectionGroupifyInformation.IsPublic, optionalParams);
            tenant.Context.ExecuteQueryRetry();
        }
        #endregion

        #region Enable Comm Site

        private static readonly Guid COMMSITEDESIGNPACKAGEID = new Guid("d604dac3-50d3-405e-9ab9-d4713cda74ef");
        /// <summary>
        /// Enable communication site on the root site of a tenant
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteUrl">Root site url of your tenant</param>
        public static void EnableCommSite(this Tenant tenant, string siteUrl = "")
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                var rootUrl = tenant.GetRootSiteUrl();
                tenant.Context.ExecuteQueryRetry();
                siteUrl = rootUrl.Value;
            }
            tenant.EnableCommSite(siteUrl, COMMSITEDESIGNPACKAGEID);
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

        #region User rights
#if !ONPREMISES
        public static Boolean IsCurrentUserTenantAdmin(ClientContext clientContext)
        {
            // Get the URL of the current site collection
            var site = clientContext.Site;
            site.EnsureProperty(s => s.Url);

            // If we are already with a context for the Admin Site, all good, the user is an admin
            if (site.Url.Contains("-admin.sharepoint.com"))
            {
                return (true);
            }
            else
            {
                // Otherwise, we need to target the Admin Site
                var siteUrl = site.Url.EndsWith("/") ? site.Url : $"{site.Url}/";
                var rootSiteUrl = siteUrl.Substring(0, siteUrl.IndexOf("/", siteUrl.IndexOf("sharepoint.com/")));
                var adminSiteUrl = rootSiteUrl.Replace(".sharepoint.com", "-admin.sharepoint.com");

                try
                {
                    // Connect to the Admin Site
                    using (var adminContext = clientContext.Clone(adminSiteUrl))
                    {
                        // Do something with the Tenant Admin Context
                        Tenant tenant = new Tenant(adminContext);
                        tenant.EnsureProperty(t => t.RootSiteUrl);

                        // If we've got access to the tenant admin context, 
                        // it means that the currently connecte user is an admin
                        return (true);
                    }
                }
                catch
                {
                    // In case of any connection exception, the user is not an admin
                    return (false);
                }
            }
        }

#elif !SP2013 && !SP2016     
        public static bool IsCurrentUserTenantAdmin(ClientContext clientContext, string tenantAdminSiteUrl)
        {
            bool result = false;
            // Get the URL of the current site collection
            var web = clientContext.Web;
            var site = clientContext.Site;
            site.EnsureProperty(s => s.Url);
            var baseTempalteId = web.GetBaseTemplateId();

            if (string.Equals(baseTempalteId, "TENANTADMIN#0", StringComparison.InvariantCultureIgnoreCase))
            {
                result = true;
            }
            else
            {
                // Otherwise, we need to target the Admin Site
                // No easy way to detect tenant admin site in on-premises, so users have to specify it
                string adminSiteUrl = tenantAdminSiteUrl;
                if (!string.IsNullOrEmpty(adminSiteUrl))
                {
                    result = CanConnectTenantAdminSite(clientContext, adminSiteUrl);
                }
                else
                {
                    //TODO: try to find a way to get the real tenant admin site url
                    var foundAdminSiteUrl = GetTenantAdminSite(clientContext);
                    if (!string.IsNullOrEmpty(foundAdminSiteUrl.AbsoluteUri))
                    {
                        result = CanConnectTenantAdminSite(clientContext, foundAdminSiteUrl.AbsoluteUri);
                    }
                    else
                    {
                        Uri uri = new Uri(clientContext.Url.TrimEnd(new[] { '/' }));
                        var rootSiteUrl = $"{uri.Scheme}://{uri.DnsSafeHost}";

                        var urlsToTry = new System.Collections.Generic.List<string>()
                        {
                            rootSiteUrl + "/sites/admin",
                            rootSiteUrl + "/sites/tenantadmin"
                        };

                        foreach (var url in urlsToTry)
                        {
                            result = CanConnectTenantAdminSite(clientContext, url);
                            if (result)
                            {
                                break;
                            }
                        }
                    }                    
                }
            }

            return result;
        }


        /// <summary>
        /// Gets the Uri for the tenant's admin site (if that one has already been created)
        /// </summary>
        /// <param name="clientContext">Context to operate against</param>
        /// <returns>The Uri holding the admin site URL</returns>
        private static Uri GetTenantAdminSite(ClientContext clientContext)
        {
            Uri uri = new Uri(clientContext.Url.TrimEnd(new[] { '/' }));
            var rootSiteUrl = $"{uri.Scheme}://{uri.DnsSafeHost}";

            // Assume there's only one admin site
            var results = clientContext.Web.SiteSearch($"contentclass:STS_Site AND SiteTemplate:TENANTADMIN AND Path:{rootSiteUrl}");
            foreach (var site in results)
            {
                return new Uri(site.Url);
            }

            return null;
        }

        private static bool CanConnectTenantAdminSite(ClientContext clientContext, string adminSiteUrl)
        {
            bool result = false;
            try
            {
                // Connect to the Admin Site
                using (var adminContext = clientContext.Clone(adminSiteUrl))
                {
                    // Do something with the Tenant Admin Context
                    Tenant tenant = new Tenant(adminContext);
                    tenant.EnsureProperty(t => t.RootSiteUrl);

                    // If we've got access to the tenant admin context, 
                    // it means that the currently connecte user is an admin
                    result = true;
                }
            }
            catch
            {
                // In case of any connection exception, the user is not an admin
                result = false;
            }

            return result;
        }
#endif
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
        /// Checks if a site collection exists, relies on tenant admin API. Sites that are recycled also return as existing sites, but with a different flag
        /// </summary>
        /// <param name="tenant">A tenant object pointing to the context of a Tenant Administration site</param>
        /// <param name="siteFullUrl">URL to the site collection</param>
        /// <returns>An enumerated type that can be: No, Yes, Recycled</returns>
        public static SiteExistence SiteExistsAnywhere(this Tenant tenant, string siteFullUrl)
        {
            try
            {
                //Get the site name
                var properties = tenant.GetSitePropertiesByUrl(siteFullUrl, false);
                tenant.Context.Load(properties);
                tenant.Context.ExecuteQueryRetry();

                // Will cause an exception if site URL is not there. Not optimal, but the way it works.
                return SiteExistence.Yes;
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
                            if (deletedProperties.Status.Equals("Recycled", StringComparison.OrdinalIgnoreCase))
                            {
                                return SiteExistence.Recycled;
                            }
                            else
                            {
                                return SiteExistence.No;
                            }
                        }
                        catch
                        {
                            return SiteExistence.No;
                        }
                    }
                    else
                    {
                        return SiteExistence.No;
                    }
                }
#if ONPREMISES
                else if (IsFileNotFoundException(ex))
                {
                    return SiteExistence.No;
                }
#endif
                else
                {
                    return SiteExistence.Yes;
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

#region Private helper methods
#if !ONPREMISES
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
#endif

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

        private static bool IsFileNotFoundException(Exception ex)
        {
            if (ex is ServerException)
            {
                if (((ServerException)ex).ServerErrorCode == -2147024894
                    && ((ServerException)ex).ServerErrorTypeName.Equals("System.IO.FileNotFoundException", StringComparison.InvariantCultureIgnoreCase))
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

#region Utilities

#if !ONPREMISES
        public static string GetTenantIdByUrl(string tenantUrl)
        {
            var tenantName = GetTenantNameFromUrl(tenantUrl);
            if (tenantName == null) return null;

            var url = $"https://login.microsoftonline.com/{tenantName}.onmicrosoft.com/.well-known/openid-configuration";
            var response = HttpHelper.MakeGetRequestForString(url);
            var json = JToken.Parse(response);

            var tokenEndpointUrl = json["token_endpoint"].ToString();
            return GetTenantIdFromAadEndpointUrl(tokenEndpointUrl);
        }
#endif

        private static string GetTenantNameFromUrl(string tenantUrl)
        {
            if (tenantUrl.ToLower().Contains("-admin.sharepoint."))
            {
                return GetSubstringFromMiddle(tenantUrl, "https://", "-admin.sharepoint.");
            }
            else
            {
                return GetSubstringFromMiddle(tenantUrl, "https://", ".sharepoint.");
            }
        }

        private static string GetTenantIdFromAadEndpointUrl(string aadEndpointUrl)
        {
            return GetSubstringFromMiddle(aadEndpointUrl, "https://login.microsoftonline.com/", "/oauth2/");
        }

        private static string GetSubstringFromMiddle(string originalString, string prefix, string suffix)
        {
            var index = originalString.IndexOf(suffix, StringComparison.OrdinalIgnoreCase);
            return index != -1 ? originalString.Substring(prefix.Length, index - prefix.Length) : null;
        }

#if !ONPREMISES
        public static string GetTenantRootSiteUrl(this Tenant tenant)
        {
            string result = null;
            tenant.EnsureProperty(t => t.RootSiteUrl);
            result = tenant.RootSiteUrl;

            /*
            var rootUrl = tenant.GetRootSiteUrl();
            tenant.Context.ExecuteQueryRetry();
            result = rootUrl.Value;
            */

            return result;
        }
#else
        public static string GetTenantRootSiteUrl(this Tenant tenant)
        {
            tenant.EnsureProperty(t => t.RootSiteUrl);

            string result = tenant.RootSiteUrl;

            if(string.IsNullOrEmpty(result))
            {
                // Onpremises (SP2019) will always return string.Emtpy for tenant.RootSiteUrl

                //tenant.EnsureProperty(t => t.RootSiteUrl);
                //var tenantUri = new Uri(tenant.Context.Url);
                //var rootSiteUri = new Uri(tenantUri.Scheme + "://" + tenantUri.Host + "/");
                //result = rootSiteUri.ToString();

                Uri uri = new Uri(tenant.Context.Url.TrimEnd(new[] { '/' }));
                result = $"{uri.Scheme}://{uri.DnsSafeHost}";
            }

            return result;
        }
#endif

#endregion

    }

    /// <summary>
    /// Defines the existence status of a Site Collection
    /// </summary>
    public enum SiteExistence
    {
        /// <summary>
        /// The Site Collection does not exist
        /// </summary>
        No,
        /// <summary>
        /// The Site Collection exists
        /// </summary>
        Yes,
        /// <summary>
        /// The Site Collection is in the Recycle Bin
        /// </summary>
        Recycled,
    }
}
