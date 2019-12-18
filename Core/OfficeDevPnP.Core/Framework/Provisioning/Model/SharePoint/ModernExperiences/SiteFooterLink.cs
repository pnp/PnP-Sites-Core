using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Footer Link for the target site
    /// </summary>
    public partial class SiteFooterLink : BaseModel, IEquatable<SiteFooterLink>
    {
        #region Public Members

        /// <summary>
        /// Defines a collection of children Footer Link for the current Footer Link (which represents an header)
        /// </summary>
        public SiteFooterLinkCollection FooterLinks { get; internal set; }

        /// <summary>
        /// Defines the DisplayName for the Footer Link for the target site
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// Defines the URL for the Footer Link for the target site
        /// </summary>
        public String Url { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for SiteFooter
        /// </summary>
        public SiteFooterLink()
        {
            this.FooterLinks = new SiteFooterLinkCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                FooterLinks.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                DisplayName.GetHashCode(),
                Url?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteFooterLink class
        /// </summary>
        /// <param name="obj">Object that represents SiteFooterLink</param>
        /// <returns>Checks whether object is SiteFooterLink class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteFooterLink))
            {
                return (false);
            }
            return (Equals((SiteFooterLink)obj));
        }

        /// <summary>
        /// Compares SiteFooterLink object based on FooterLinks, DisplayName, and Url
        /// </summary>
        /// <param name="other">SiteFooterLink Class object</param>
        /// <returns>true if the SiteFooterLink object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteFooterLink other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.FooterLinks.DeepEquals(other.FooterLinks) &&
                this.DisplayName == other.DisplayName &&
                this.Url == other.Url
                );
        }

        #endregion
    }
}
