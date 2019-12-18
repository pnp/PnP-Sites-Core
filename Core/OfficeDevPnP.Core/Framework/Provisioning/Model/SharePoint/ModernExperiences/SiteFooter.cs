using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Footer settings for the target site
    /// </summary>
    public partial class SiteFooter : BaseModel, IEquatable<SiteFooter>
    {
        #region Public Members

        /// <summary>
        /// Defines whether the site Footer is enabled or not
        /// </summary>
        public Boolean Enabled { get; set; }

        /// <summary>
        /// Defines the Logo to render in the Footer
        /// </summary>
        public String Logo { get; set; }

        /// <summary>
        /// Defiones the name of the footer. Only visible if "NameVisiblity" has been set to true.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Defines whether the existing site Footer links should be removed
        /// </summary>
        public Boolean RemoveExistingNodes { get; set; }

        /// <summary>
        /// Defines the Footer Links for the target site
        /// </summary>
        public SiteFooterLinkCollection FooterLinks { get; private set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for SiteFooter
        /// </summary>
        public SiteFooter()
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
            return (String.Format("{0}|{1}|{2}|{3}|{4}|",
                Enabled.GetHashCode(),
                Logo?.GetHashCode() ?? 0,
                Name?.GetHashCode(),
                RemoveExistingNodes.GetHashCode(),
                FooterLinks.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteFooter class
        /// </summary>
        /// <param name="obj">Object that represents SiteFooter</param>
        /// <returns>Checks whether object is SiteFooter class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteFooter))
            {
                return (false);
            }
            return (Equals((SiteFooter)obj));
        }

        /// <summary>
        /// Compares SiteFooter object based on Layout, MenuStyle, and BackgroundEmphasis
        /// </summary>
        /// <param name="other">SiteFooter Class object</param>
        /// <returns>true if the SiteFooter object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteFooter other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Enabled == other.Enabled &&
                this.Logo == other.Logo &&
                this.Name == other.Name &&
                this.RemoveExistingNodes == other.RemoveExistingNodes &&
                this.FooterLinks.DeepEquals(other.FooterLinks)
                );
        }

        #endregion
    }
}
