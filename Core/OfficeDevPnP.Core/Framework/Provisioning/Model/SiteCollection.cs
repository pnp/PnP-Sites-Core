using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class SiteCollection: BaseProvisioningModel, IEquatable<SiteCollection>
    {
        #region Private Members

        #endregion

        #region Constructor

        public SiteCollection()
        {
            this.Templates = new ProvisioningTemplateCollection(this.ParentProvisioning);
            this.Sites = new SubSiteCollection(this.ParentProvisioning);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// The template to use while creating the Site Collection
        /// </summary>
        public SiteCollectionSiteTemplate SiteTemplate { get; set; }

        /// <summary>
        /// The list of Owners of the Site Collection
        /// </summary>
        public String Owners { get; set; }

        /// <summary>
        /// Declares whether the current Site Collection is the Hub Site of a new Hub
        /// </summary>
        public Boolean IsHubSite { get; set; }

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
        /// Defines the list of Provisioning Templates to apply to the site collection, if any
        /// </summary>
        public ProvisioningTemplateCollection Templates { get; private set; }

        /// <summary>
        /// Defines the list of sub-sites, if any
        /// </summary>
        public SubSiteCollection Sites { get; private set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|",
                this.SiteTemplate.GetHashCode(),
                this.Owners.GetHashCode(),
                this.IsHubSite.GetHashCode(),
                this.Title.GetHashCode(),
                this.Alias.GetHashCode(),
                this.Language.GetHashCode(),
                this.Templates.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Sites.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteCollection
        /// </summary>
        /// <param name="obj">Object that represents SiteCollection</param>
        /// <returns>true if the current object is equal to the SiteCollection</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteCollection))
            {
                return (false);
            }
            return (Equals((SiteCollection)obj));
        }

        /// <summary>
        /// Compares SiteCollection object based on its properties
        /// </summary>
        /// <param name="other">SiteCollection object</param>
        /// <returns>true if the SiteCollection object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteCollection other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.SiteTemplate == other.SiteTemplate &&
                this.Owners == other.Owners &&
                this.IsHubSite== other.IsHubSite &&
                this.Title == other.Title &&
                this.Alias == other.Alias &&
                this.Language == other.Language &&
                this.Templates.Intersect(other.Templates).Count() == 0 &&
                this.Sites.DeepEquals(other.Sites)
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the template for a new "modern" Site Collection
    /// </summary>
    public enum SiteCollectionSiteTemplate
    {
        /// <summary>
        /// A Communication Site
        /// </summary>
        Communication,
        /// <summary>
        /// A "modern" Team Site
        /// </summary>
        Team,
        /// <summary>
        /// A "modern" Team Site without the corresponding Office 365 Group
        /// </summary>
        TeamNoGroup,
    }
}
