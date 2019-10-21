using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// The Messaging Settings for the Team
    /// </summary>
    public partial class TeamMessagingSettings : BaseModel, IEquatable<TeamMessagingSettings>
    {
        #region Public Members

        /// <summary>
        /// Defines if users can edit their messages
        /// </summary>
        public Boolean AllowUserEditMessages { get; set; }

        /// <summary>
        /// Defines if users can delete their messages
        /// </summary>
        public Boolean AllowUserDeleteMessages { get; set; }

        /// <summary>
        /// Defines if owners can delete any message
        /// </summary>
        public Boolean AllowOwnerDeleteMessages { get; set; }

        /// <summary>
        /// Defines if @team mentions are allowed
        /// </summary>
        public Boolean AllowTeamMentions { get; set; }

        /// <summary>
        /// Defines if @channel mentions are allowed
        /// </summary>
        public Boolean AllowChannelMentions { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|",
                AllowUserEditMessages.GetHashCode(),
                AllowUserDeleteMessages.GetHashCode(),
                AllowOwnerDeleteMessages.GetHashCode(),
                AllowTeamMentions.GetHashCode(),
                AllowChannelMentions.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamMessagingSettings class
        /// </summary>
        /// <param name="obj">Object that represents TeamMessagingSettings</param>
        /// <returns>Checks whether object is TeamMessagingSettings class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamMessagingSettings))
            {
                return (false);
            }
            return (Equals((TeamMessagingSettings)obj));
        }

        /// <summary>
        /// Compares TeamMessagingSettings object based on AllowUserEditMessages, AllowUserDeleteMessages, AllowOwnerDeleteMessages, AllowTeamMentions, and AllowChannelMentions
        /// </summary>
        /// <param name="other">TeamMessagingSettings Class object</param>
        /// <returns>true if the TeamMessagingSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamMessagingSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AllowUserEditMessages == other.AllowUserEditMessages &&
                this.AllowUserDeleteMessages == other.AllowUserDeleteMessages &&
                this.AllowOwnerDeleteMessages == other.AllowOwnerDeleteMessages &&
                this.AllowTeamMentions == other.AllowTeamMentions &&
                this.AllowChannelMentions == other.AllowChannelMentions
                );
        }

        #endregion
    }
}
