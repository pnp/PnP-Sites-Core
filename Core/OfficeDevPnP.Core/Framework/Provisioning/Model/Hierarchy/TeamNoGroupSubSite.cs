using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Concrete type defining a Team sub-Site with no Office 365 Group
    /// </summary>
    public partial class TeamNoGroupSubSite : SubSite
    {
        /// <summary>
        /// The URL of the target Site
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Language of the target Site
        /// </summary>
        public int Language { get; set; }

        /// <summary>
        /// The TimeZone of the target Site
        /// </summary>
        public int TimeZoneId { get; set; }

        protected override bool EqualsInherited(SubSite other)
        {
            if (!(other is TeamNoGroupSubSite otherTyped))
            {
                return (false);
            }

            return (this.Url == otherTyped.Url &&
                this.Language == otherTyped.Language &&
                this.TimeZoneId == otherTyped.TimeZoneId
                );
        }

        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.Url?.GetHashCode() ?? 0,
                this.Language.GetHashCode(),
                this.TimeZoneId.GetHashCode()
            ).GetHashCode());
        }
    }
}
