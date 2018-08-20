using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class SubSite : BaseProvisioningModel, IEquatable<SubSite>
    {
        #region Private Members

        #endregion

        #region Constructors

        public SubSite()
        {
            this.Templates = new ProvisioningTemplateCollection(this.ParentProvisioning);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// The template to use while creating the SubSite
        /// </summary>
        public SubSiteTemplate SiteTemplate { get; set; }

        /// <summary>
        /// Defines if the Quick launch enabled
        /// </summary>
        public Boolean QuickLaunchEnabled { get; set; }

        /// <summary>
        /// Defines whether to use the same permissions of the parent site or not
        /// </summary>
        public Boolean UseSamePermissionsAsParentSite { get; set; }

        /// <summary>
        /// Title of the site
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Defines the alias for the Office 365 Group created with the Site Collection, when needed.
        /// </summary>
        public String Alias { get; set; }

        /// <summary>
        /// Language of the target Site
        /// </summary>
        public String Language { get; set; }

        /// <summary>
        /// Defines the list of Provisioning Templates to apply to the sub-site, if any
        /// </summary>
        public ProvisioningTemplateCollection Templates { get; private set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|",
                this.SiteTemplate.GetHashCode(),
                this.QuickLaunchEnabled.GetHashCode(),
                this.UseSamePermissionsAsParentSite.GetHashCode(),
                this.Title.GetHashCode(),
                this.Alias.GetHashCode(),
                this.Language.GetHashCode(),
                this.Templates.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SubSite
        /// </summary>
        /// <param name="obj">Object that represents SubSite</param>
        /// <returns>true if the current object is equal to the SubSite</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SubSite))
            {
                return (false);
            }
            return (Equals((SubSite)obj));
        }

        /// <summary>
        /// Compares SubSite object based on its properties
        /// </summary>
        /// <param name="other">SubSite object</param>
        /// <returns>true if the SubSite object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SubSite other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.SiteTemplate == other.SiteTemplate &&
                this.QuickLaunchEnabled == other.QuickLaunchEnabled &&
                this.UseSamePermissionsAsParentSite == other.UseSamePermissionsAsParentSite &&
                this.Title == other.Title &&
                this.Alias == other.Alias &&
                this.Language == other.Language &&
                this.Templates.Intersect(other.Templates).Count() == 0
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the template for a new "modern" SubSite
    /// </summary>
    public enum SubSiteTemplate
    {
        /// <summary>
        /// A "modern" Team Site without the corresponding Office 365 Group
        /// </summary>
        TeamNoGroup,
    }
}
