using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Concrete type defining a Communication Site Collection
    /// </summary>
    public partial class CommunicationSiteCollection: SiteCollection
    {
        /// <summary>
        /// The URL of the target Site
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Owner of the target Site
        /// </summary>
        /// <remarks>
        /// Reserved for future use
        /// </remarks>
        public string Owner { get; set; }

        /// <summary>
        /// The ID of the SiteDesign
        /// </summary>
        public string SiteDesign { get; set; }

        /// <summary>
        /// Defines whether the target Site can be shared to guest users or not
        /// </summary>
        public bool AllowFileSharingForGuestUsers { get; set; }

        /// <summary>
        /// The Classification of the target Site
        /// </summary>
        public string Classification { get; set; }

        /// <summary>
        /// Language of the target Site
        /// </summary>
        public int Language { get; set; }

        protected override bool EqualsInherited(SiteCollection other)
        {
            if (!(other is CommunicationSiteCollection otherTyped))
            {
                return (false);
            }

            return (this.Url == otherTyped.Url &&
                this.Owner == otherTyped.Owner &&
                this.SiteDesign == otherTyped.SiteDesign &&
                this.AllowFileSharingForGuestUsers == otherTyped.AllowFileSharingForGuestUsers &&
                this.Classification == otherTyped.Classification &&
                this.Language == otherTyped.Language
                );
        }

        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|",
                this.Url?.GetHashCode() ?? 0,
                this.Owner?.GetHashCode() ?? 0,
                this.SiteDesign?.GetHashCode() ?? 0,
                this.AllowFileSharingForGuestUsers.GetHashCode(),
                this.Classification?.GetHashCode() ?? 0,
                this.Language.GetHashCode()
            ).GetHashCode());
        }
    }
}
