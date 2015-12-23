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
    public partial class WebSettings : IEquatable<WebSettings>
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

        #endregion

        #region Constructors

        public WebSettings() { }

        public WebSettings(Boolean noCrawl, String requestAccessEmail, String welcomePage)
        {
            this.NoCrawl = noCrawl;
            this.RequestAccessEmail = requestAccessEmail;
            this.WelcomePage = welcomePage;
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                (this.NoCrawl.GetHashCode()),
                (this.RequestAccessEmail != null ? this.RequestAccessEmail.GetHashCode() : 0),
                (this.WelcomePage != null ? this.WelcomePage.GetHashCode() : 0)
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
            return (this.NoCrawl == other.NoCrawl &&
                    this.RequestAccessEmail == other.RequestAccessEmail &&
                    this.WelcomePage == other.WelcomePage
                );
        }

        #endregion
    }
}
