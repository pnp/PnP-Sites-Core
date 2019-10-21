using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines the Configuration for a Team Tab
    /// </summary>
    public partial class TeamTabConfiguration : BaseModel, IEquatable<TeamTabConfiguration>
    {
        #region Public Members

        /// <summary>
        /// Identifier for the entity hosted by the Tab provider
        /// </summary>
        public String EntityId { get; set; }

        /// <summary>
        /// Url used for rendering Tab contents in Teams
        /// </summary>
        public String ContentUrl { get; set; }


        /// <summary>
        /// Url called by Teams client when a Tab is removed using the Teams Client
        /// </summary>
        public String RemoveUrl { get; set; }

        /// <summary>
        /// Url for showing Tab contents outside of Teams
        /// </summary>
        public String WebsiteUrl { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                EntityId?.GetHashCode() ?? 0,
                ContentUrl?.GetHashCode() ?? 0,
                RemoveUrl?.GetHashCode() ?? 0,
                WebsiteUrl?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamTabConfiguration class
        /// </summary>
        /// <param name="obj">Object that represents TeamTabConfiguration</param>
        /// <returns>Checks whether object is TeamTabConfiguration class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamTabConfiguration))
            {
                return (false);
            }
            return (Equals((TeamTabConfiguration)obj));
        }

        /// <summary>
        /// Compares TeamTabConfiguration object based on EntityId, ContentUrl, RemoveUrl, and WebsiteUrl
        /// </summary>
        /// <param name="other">TeamTabConfiguration Class object</param>
        /// <returns>true if the TeamTabConfiguration object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamTabConfiguration other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.EntityId == other.EntityId &&
                this.ContentUrl == other.ContentUrl &&
                this.RemoveUrl == other.RemoveUrl &&
                this.WebsiteUrl == other.WebsiteUrl
                );
        }

        #endregion
    }
}
