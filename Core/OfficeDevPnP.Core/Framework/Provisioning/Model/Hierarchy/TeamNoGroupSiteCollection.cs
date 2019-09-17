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

        protected override bool EqualsInherited(SiteCollection other)
        {
            if (!(other is TeamNoGroupSiteCollection otherTyped))
            {
                return (false);
            }

            return (this.Url == otherTyped.Url &&
                this.Owner == otherTyped.Owner &&
                this.Language == otherTyped.Language &&
                this.TimeZoneId == otherTyped.TimeZoneId
                );
        }

        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                this.Url?.GetHashCode() ?? 0,
                this.Owner?.GetHashCode() ?? 0,
                this.Language.GetHashCode(),
                this.TimeZoneId.GetHashCode()
            ).GetHashCode());
        }
    }
}
