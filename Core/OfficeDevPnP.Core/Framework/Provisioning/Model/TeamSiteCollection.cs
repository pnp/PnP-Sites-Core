using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Concrete type defining a Team Site Collection
    /// </summary>
    public partial class TeamSiteCollection : SiteCollection
    {
        /// <summary>
        /// The Alias of the target Site
        /// </summary>
        public string Alias { get; set; }

        /// <summary>
        /// The DisplayName of the target Site
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Defines whether the Office 365 Group associated with the Site is Public or Private
        /// </summary>
        public bool IsPublic { get; set; }

        /// <summary>
        /// The Classification of the target Site
        /// </summary>
        public string Classification { get; set; }

        protected override bool EqualsInherited(SiteCollection other)
        {
            if (!(other is TeamSiteCollection otherTyped))
            {
                return (false);
            }

            return (this.Alias == otherTyped.Alias &&
                this.DisplayName == otherTyped.DisplayName &&
                this.IsPublic == otherTyped.IsPublic &&
                this.Classification == otherTyped.Classification
                );
        }

        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                this.Alias.GetHashCode(),
                this.DisplayName.GetHashCode(),
                this.IsPublic.GetHashCode(),
                this.Classification.GetHashCode()
            ).GetHashCode());
        }
    }
}
