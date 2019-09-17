using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// The Guest Settings for the Team
    /// </summary>
    public partial class TeamGuestSettings : BaseModel, IEquatable<TeamGuestSettings>
    {
        #region Public Members

        /// <summary>
        /// Defines whether Guests are allowed to create Channels or not
        /// </summary>
        public Boolean AllowCreateUpdateChannels { get; set; }

        /// <summary>
        /// Defines whether Guests are allowed to delete Channels or not
        /// </summary>
        public Boolean AllowDeleteChannels { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                AllowCreateUpdateChannels.GetHashCode(),
                AllowDeleteChannels.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamGuestSettings class
        /// </summary>
        /// <param name="obj">Object that represents TeamGuestSettings</param>
        /// <returns>Checks whether object is TeamGuestSettings class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamGuestSettings))
            {
                return (false);
            }
            return (Equals((TeamGuestSettings)obj));
        }

        /// <summary>
        /// Compares TeamGuestSettings object based on AllowGiphy, GiphyContentRating, AllowStickersAndMemes, and AllowCustomMemes
        /// </summary>
        /// <param name="other">TeamGuestSettings Class object</param>
        /// <returns>true if the TeamGuestSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamGuestSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AllowCreateUpdateChannels == other.AllowCreateUpdateChannels &&
                this.AllowDeleteChannels == other.AllowDeleteChannels
                );
        }

        #endregion
    }
}
