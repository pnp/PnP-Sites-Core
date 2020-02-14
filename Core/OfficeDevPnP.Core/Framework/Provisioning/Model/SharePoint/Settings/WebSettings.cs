using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object used in the Provisioning template that defines a Section of Settings for the current Web Site
    /// </summary>
    public partial class WebSettings : BaseModel, IEquatable<WebSettings>
    {
        #region Properties

        /// <summary>
        /// Defines whether the site has to be crawled or not
        /// </summary>
        public Boolean NoCrawl { get; set; }

        /// <summary>
        /// The email address to which any access request will be sent
        /// </summary>
        public String RequestAccessEmail { get; set; }

        /// <summary>
        /// Defines the Welcome Page (Home Page) of the site to which the Provisioning Template is applied.
        /// </summary>
        public String WelcomePage { get; set; }

        /// <summary>
        /// The Title of the Site, optional attribute.
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// The Description of the Site, optional attribute.
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The SiteLogo of the Site, optional attribute.
        /// </summary>
        public String SiteLogo { get; set; }

        /// <summary>
        /// The AlternateCSS of the Site, optional attribute.
        /// </summary>
        public String AlternateCSS { get; set; }

        /// <summary>
        /// The MasterPage Url of the Site, optional attribute.
        /// </summary>
        public String MasterPageUrl { get; set; }

        /// <summary>
        /// The Custom MasterPage Url of the Site, optional attribute.
        /// </summary>
        public String CustomMasterPageUrl { get; set; }

        /// <summary>
        /// The Hub Site Url of the Site, optional attribute.
        /// </summary>
        public String HubSiteUrl { get; set; }

        /// <summary>
        /// Defines whether the comments on site pages are disabled or not
        /// </summary>
        public Boolean CommentsOnSitePagesDisabled { get; set; }

        /// <summary>
        /// Enables or disables the QuickLaunch for the site
        /// </summary>
        public Boolean QuickLaunchEnabled { get; set; }

        /// <summary>
        /// Defines the list of Alternate UI Cultures for the current web
        /// </summary>
        public AlternateUICultureCollection AlternateUICultures { get; set; }

        /// <summary>
        /// Defines whether to enable Multilingual capabilities for the current web
        /// </summary>
        public bool IsMultilingual { get; set; }

        /// <summary>
        /// Defines whether to OverwriteTranslationsOnChange on change for the current web
        /// </summary>
        public bool OverwriteTranslationsOnChange { get; set; }

        /// <summary>
        /// Defines whether to exclude the web from offline client
        /// </summary>
        public bool ExcludeFromOfflineClient { get; set; }

        /// <summary>
        /// Defines whether members can share content from the current web
        /// </summary>
        public bool MembersCanShare { get; set; }

        /// <summary>
        /// Defines whether disable flows for the current web
        /// </summary>
        public bool DisableFlows { get; set; }

        /// <summary>
        /// Defines whether disable PowerApps for the current web
        /// </summary>
        public bool DisableAppViews { get; set; }

        /// <summary>
        /// Defines whether to enable the Horizontal QuickLaunch for the current web
        /// </summary>
        public bool HorizontalQuickLaunch { get; set; }

        /// <summary>
        /// Defines the SearchScope for the site
        /// </summary>
        public SearchScopes SearchScope { get; set; }

        /// <summary>
        /// Define if the suitebar search box should show or not 
        /// </summary>
        public SearchBoxInNavBar SearchBoxInNavBar { get; set; }

        /// <summary>
        /// Defines the Search Center URL
        /// </summary>
        public string SearchCenterUrl { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Default Constructor
        /// </summary>
        public WebSettings()
        {
            this.AlternateUICultures = new AlternateUICultureCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="noCrawl">Based on boolean values sets crawl to the site or subsite</param>
        /// <param name="requestAccessEmail">E-mail address for request access</param>
        /// <param name="welcomePage">Welcome page for site or subsite</param>
        public WebSettings(Boolean noCrawl, String requestAccessEmail, String welcomePage) :
            this(noCrawl, requestAccessEmail, welcomePage, null, null, null, null)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="noCrawl">Based on boolean values sets crawl to the site or subsite</param>
        /// <param name="requestAccessEmail">E-mail address for request access</param>
        /// <param name="welcomePage">Welcome page for site or subsite</param>
        /// <param name="title">Title of site or subsite</param>
        /// <param name="description">Description of site or subsite</param>
        /// <param name="siteLogo">Logo of site or subsite</param>
        /// <param name="alternateCSS">Alternate css file location of site or subsite</param>
        public WebSettings(Boolean noCrawl, String requestAccessEmail, String welcomePage,
            String title, String description, String siteLogo, String alternateCSS) : this()
        {
            this.NoCrawl = noCrawl;
            this.RequestAccessEmail = requestAccessEmail;
            this.WelcomePage = welcomePage;
            this.Title = title;
            this.Description = description;
            this.SiteLogo = siteLogo;
            this.AlternateCSS = alternateCSS;
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets hash code
        /// </summary>
        /// <returns>Returns hash code in integer</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}",
                (this.NoCrawl.GetHashCode()),
                (this.RequestAccessEmail != null ? this.RequestAccessEmail.GetHashCode() : 0),
                (this.WelcomePage != null ? this.WelcomePage.GetHashCode() : 0),
                (this.Title != null ? this.Title.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.SiteLogo != null ? this.SiteLogo.GetHashCode() : 0),
                (this.AlternateCSS != null ? this.AlternateCSS.GetHashCode() : 0),
                (this.HubSiteUrl != null ? this.HubSiteUrl.GetHashCode() : 0),
                this.CommentsOnSitePagesDisabled.GetHashCode(),
                this.QuickLaunchEnabled.GetHashCode(),
                AlternateUICultures.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.IsMultilingual.GetHashCode(),
                this.OverwriteTranslationsOnChange.GetHashCode(),
                this.ExcludeFromOfflineClient.GetHashCode(),
                this.MembersCanShare.GetHashCode(),
                this.DisableFlows.GetHashCode(),
                this.DisableAppViews.GetHashCode(),
                this.HorizontalQuickLaunch.GetHashCode(),
                this.SearchScope.GetHashCode(),
                this.SearchBoxInNavBar.GetHashCode(),
                this.SearchCenterUrl.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares web settings with other web settings
        /// </summary>
        /// <param name="obj">WebSettings object</param>
        /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is WebSettings))
            {
                return (false);
            }
            return (Equals((WebSettings)obj));
        }

        /// <summary>
        /// Compares web settings with other web settings
        /// </summary>
        /// <param name="other">WebSettings object</param>
        /// <returns>true if the WebSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(WebSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.NoCrawl == other.NoCrawl &&
                    this.RequestAccessEmail == other.RequestAccessEmail &&
                    this.WelcomePage == other.WelcomePage &&
                    this.Title == other.Title &&
                    this.Description == other.Description &&
                    this.SiteLogo == other.SiteLogo &&
                    this.AlternateCSS == other.AlternateCSS &&
                    this.HubSiteUrl == other.HubSiteUrl &&
                    this.CommentsOnSitePagesDisabled == other.CommentsOnSitePagesDisabled &&
                    this.QuickLaunchEnabled == other.QuickLaunchEnabled &&
                    this.AlternateUICultures.DeepEquals(other.AlternateUICultures) &&
                    this.IsMultilingual == other.IsMultilingual &&
                    this.OverwriteTranslationsOnChange == other.OverwriteTranslationsOnChange &&
                    this.ExcludeFromOfflineClient == other.ExcludeFromOfflineClient &&
                    this.MembersCanShare == other.MembersCanShare &&
                    this.DisableFlows == other.DisableFlows &&
                    this.DisableAppViews == other.DisableAppViews &&
                    this.HorizontalQuickLaunch == other.HorizontalQuickLaunch &&
                    this.SearchScope == other.SearchScope &&
                    this.SearchBoxInNavBar == other.SearchBoxInNavBar &&
                    this.SearchCenterUrl == other.SearchCenterUrl
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the SearchScope for the site
    /// </summary>
    public enum SearchScopes
    {
        /// <summary>
        /// Defines the DefaultScope for the SearchScope of the site
        /// </summary>
        DefaultScope,
        /// <summary>
        /// Defines the Tenant for the SearchScope of the site
        /// </summary>
        Tenant,
        /// <summary>
        /// Defines the Hub for the SearchScope of the site
        /// </summary>
        Hub,
        /// <summary>
        /// Defines the Site for the SearchScope of the site
        /// </summary>
        Site,
    }

    public enum SearchBoxInNavBar
    {
        Inherit = 0,
        AllPages = 1,
        ModernOnly = 2,
        Hidden = 3
    }
}
