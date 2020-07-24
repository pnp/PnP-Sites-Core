using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Concrete type defining a Team Site Collection without an Office 365 Group
    /// </summary>
    public partial class TeamNoGroupSiteCollection : SiteCollection
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
        /// Language of the target Site
        /// </summary>
        public int Language { get; set; }

        /// <summary>
        /// The TimeZone of the target Site
        /// </summary>
        public int TimeZoneId { get; set; }

        /// <summary>
        /// Declare whether to groupify the team site after creation or not
        /// </summary>
        public bool Groupify { get; set; }

        /// <summary>
        /// The Alias of the target Office 365 Group backing the Site, optional attribute. It is used if and only if the Groupify attribute has a value of True.
        /// </summary>
        public string Alias { get; set; }

        /// <summary>
        /// The Classification of the target groupified Site, if any, optional attribute. It is used if and only if the Groupify attribute has a value of True.
        /// </summary>
        public string Classification { get; set; }

        /// <summary>
        /// Defines whether the Office 365 Group for the target groupified Site is Public or Private, optional attribute. It is used if and only if the Groupify attribute has a value of True.
        /// </summary>
        public bool IsPublic { get; set; }

        /// <summary>
        /// Defines whether to keep the old home page of the site after it gets groupified, optional attribute. It is used if and only if the Groupify attribute has a value of True.
        /// </summary>
        public bool KeepOldHomePage { get; set; }

        protected override bool EqualsInherited(SiteCollection other)
        {
            if (!(other is TeamNoGroupSiteCollection otherTyped))
            {
                return (false);
            }

            return (this.Url == otherTyped.Url &&
                this.Owner == otherTyped.Owner &&
                this.Language == otherTyped.Language &&
                this.TimeZoneId == otherTyped.TimeZoneId &&
                this.Groupify == otherTyped.Groupify &&
                this.Alias == otherTyped.Alias &&
                this.Classification == otherTyped.Classification &&
                this.IsPublic == otherTyped.IsPublic &&
                this.KeepOldHomePage == otherTyped.KeepOldHomePage
                );
        }

        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|",
                this.Url?.GetHashCode() ?? 0,
                this.Owner?.GetHashCode() ?? 0,
                this.Language.GetHashCode(),
                this.TimeZoneId.GetHashCode(),
                this.Groupify.GetHashCode(),
                this.Alias.GetHashCode(),
                this.Classification.GetHashCode(),
                this.IsPublic.GetHashCode(),
                this.KeepOldHomePage.GetHashCode()
            ).GetHashCode());
        }
    }
}
