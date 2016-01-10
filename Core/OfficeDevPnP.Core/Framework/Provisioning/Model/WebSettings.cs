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

        #endregion

        #region Constructors

        public WebSettings() { }

        public WebSettings(Boolean noCrawl, String requestAccessEmail, String welcomePage):
            this(noCrawl, requestAccessEmail, welcomePage, null, null, null, null)
        {
        }

        public WebSettings(Boolean noCrawl, String requestAccessEmail, String welcomePage, 
            String title, String description, String siteLogo, String alternateCSS)
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

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                (this.NoCrawl.GetHashCode()),
                (this.RequestAccessEmail != null ? this.RequestAccessEmail.GetHashCode() : 0),
                (this.WelcomePage != null ? this.WelcomePage.GetHashCode() : 0),
                (this.Title != null ? this.Title.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.SiteLogo != null ? this.SiteLogo.GetHashCode() : 0),
                (this.AlternateCSS != null ? this.AlternateCSS.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is WebSettings))
            {
                return (false);
            }
            return (Equals((WebSettings)obj));
        }

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
                    this.AlternateCSS == other.AlternateCSS
                );
        }

        #endregion
    }
}
