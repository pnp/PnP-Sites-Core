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

        /// <summary>
        /// Defines whether to create a Microsoft Team backing the modern Team Site
        /// </summary>
        public bool Teamify { get; set; }

        /// <summary>
        /// Defines whether to hide the create a Microsoft Team option in the UI of the Team Site
        /// </summary>
        public bool HideTeamify { get; set; }
        
        /// <summary>
        /// Language of the target Site
        /// </summary>
        public int Language { get; set; }

        /// <summary>
        /// Allows to associate the Office 365 Group associated with the Team Site to a Group Lifecycle Policy
        /// </summary>
        public string GroupLifecyclePolicyId { get; set; }

        /// <summary>
        /// The ID of the SiteDesign, if any, to apply to the target Site
        /// </summary>
        public string SiteDesign { get; set; }

        protected override bool EqualsInherited(SiteCollection other)
        {
            if (!(other is TeamSiteCollection otherTyped))
            {
                return (false);
            }

            return (this.Alias == otherTyped.Alias &&
                this.DisplayName == otherTyped.DisplayName &&
                this.IsPublic == otherTyped.IsPublic &&
                this.Classification == otherTyped.Classification &&
                this.Teamify == otherTyped.Teamify &&
                this.HideTeamify == otherTyped.HideTeamify &&
                this.GroupLifecyclePolicyId == otherTyped.GroupLifecyclePolicyId &&
                this.Language == otherTyped.Language &&
                this.SiteDesign == otherTyped.SiteDesign
                );
        }

        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|",
                this.Alias?.GetHashCode() ?? 0,
                this.DisplayName?.GetHashCode() ?? 0,
                this.IsPublic.GetHashCode(),
                this.Classification?.GetHashCode() ?? 0,
                this.Teamify.GetHashCode(),
                this.HideTeamify.GetHashCode(),
                this.GroupLifecyclePolicyId?.GetHashCode() ?? 0,
                this.Language.GetHashCode(),
                this.SiteDesign.GetHashCode()
            ).GetHashCode());
        }
    }
}
