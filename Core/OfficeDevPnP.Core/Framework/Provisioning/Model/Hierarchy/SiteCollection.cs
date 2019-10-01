using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public abstract partial class SiteCollection : BaseHierarchyModel, IEquatable<SiteCollection>
    {
        #region Private Members
        Guid id;
        #endregion

        #region Constructor

        public SiteCollection()
        {
            this.Templates = new List<String>();
            this.Sites = new SubSiteCollection(this.ParentHierarchy);
            this.id = Guid.NewGuid();
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Declares whether the current Site Collection is the Hub Site of a new Hub
        /// </summary>
        public Boolean IsHubSite { get; set; }

        /// <summary>
        /// Defines the url to the logo if this site is a hubsite. Only applicable if IsHubSite is set to true.
        /// </summary>
        public string HubSiteLogoUrl { get; set; }

        /// <summary>
        /// Defines the url to the logo if this site is a hubsite. Only applicable if IsHubSite is set to true.
        /// </summary>
        public string HubSiteTitle { get; set; }

        /// <summary>
        /// Title of the site
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Description of the site
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// Defines the list of Provisioning Templates to apply to the site collection, if any
        /// </summary>
        public List<String> Templates { get; internal set; }

        /// <summary>
        /// Defines the list of sub-sites, if any
        /// </summary>
        public SubSiteCollection Sites { get; private set; }

        /// <summary>
        /// Defines the Theme to apply to the SiteCollection
        /// </summary>
        public String Theme { get; set; }

        /// <summary>
        /// Internal use only
        /// </summary>
        public Guid Id => id;

        /// <summary>
        /// Defines an optional ID in the sequence for use in tokens.
        /// </summary>
        public string ProvisioningId { get; set; }

        public override string ToString()
        {
            return id.ToString();
        }
        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|",
                this.IsHubSite.GetHashCode(),
                this.Title.GetHashCode(),
                this.Description.GetHashCode(),
                this.Templates.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Sites.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Theme.GetHashCode(),
                this.ProvisioningId.GetHashCode(),
                this.GetInheritedHashCode(),
                this.HubSiteLogoUrl?.GetHashCode(),
                this.HubSiteTitle?.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Returns the HashCode of the members of any inherited type
        /// </summary>
        /// <returns></returns>
        protected abstract int GetInheritedHashCode();

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

            return (this.IsHubSite == other.IsHubSite &&
                this.Title == other.Title &&
                this.Description == other.Description &&
                this.Templates.Intersect(other.Templates).Count() == 0 &&
                this.Sites.DeepEquals(other.Sites) &&
                this.Theme == other.Theme &&
                this.ProvisioningId == other.ProvisioningId &&
                this.EqualsInherited(other) &&
                this.HubSiteLogoUrl == other.HubSiteLogoUrl &&
                this.HubSiteTitle == other.HubSiteTitle
                );
        }

        /// <summary>
        /// Compares the HashCode of the members of any inherited type
        /// </summary>
        /// <returns></returns>
        protected abstract bool EqualsInherited(SiteCollection other);

        #endregion
    }
}
