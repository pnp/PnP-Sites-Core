using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines an TeamApp for automated provisiong of Microsoft Teams
    /// </summary>
    public partial class TeamChannelMessage : BaseModel, IEquatable<TeamChannelMessage>
    {
        #region Public Members

        /// <summary>
        /// Defines a Message for a Channel in a Team
        /// </summary>
        public String Message { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                Message?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamChannelMessage class
        /// </summary>
        /// <param name="obj">Object that represents TeamChannelMessage</param>
        /// <returns>Checks whether object is TeamChannelMessage class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamChannelMessage))
            {
                return (false);
            }
            return (Equals((TeamChannelMessage)obj));
        }

        /// <summary>
        /// Compares TeamChannelMessage object based on Message
        /// </summary>
        /// <param name="other">TeamChannelMessage Class object</param>
        /// <returns>true if the TeamChannelMessage object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamChannelMessage other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Message == other.Message);
        }

        #endregion
    }
}
