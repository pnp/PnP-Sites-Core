using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// The Members Settings for the Team
    /// </summary>
    public partial class TeamMemberSettings : BaseModel, IEquatable<TeamMemberSettings>
    {
        #region Public Members

        /// <summary>
        /// Defines if members can add and update channels
        /// </summary>
        public Boolean AllowCreateUpdateChannels { get; set; }

        /// <summary>
        /// Defines if members can delete channels
        /// </summary>
        public Boolean AllowDeleteChannels { get; set; }

        /// <summary>
        /// Defines if members can add and remove apps
        /// </summary>
        public Boolean AllowAddRemoveApps { get; set; }

        /// <summary>
        /// Defines if members can add, update, and remove tabs
        /// </summary>
        public Boolean AllowCreateUpdateRemoveTabs { get; set; }

        /// <summary>
        /// Defines if members can add, update, and remove connectors
        /// </summary>
        public Boolean AllowCreateUpdateRemoveConnectors { get; set; }

        /// <summary>
        /// Defines if members can create private channels
        /// </summary>
        public bool AllowCreatePrivateChannels { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|",
                AllowCreateUpdateChannels.GetHashCode(),
                AllowDeleteChannels.GetHashCode(),
                AllowAddRemoveApps.GetHashCode(),
                AllowCreateUpdateRemoveTabs.GetHashCode(),
                AllowCreateUpdateRemoveConnectors.GetHashCode(),
                AllowCreatePrivateChannels.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamMemberSettings class
        /// </summary>
        /// <param name="obj">Object that represents TeamMemberSettings</param>
        /// <returns>Checks whether object is TeamMemberSettings class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamMemberSettings))
            {
                return (false);
            }
            return (Equals((TeamMemberSettings)obj));
        }

        /// <summary>
        /// Compares TeamFunSettings object based on AllowCreateUpdateChannels, AllowDeleteChannels, AllowAddRemoveApps, AllowCreateUpdateRemoveTabs, and AllowCreateUpdateRemoveConnectors
        /// </summary>
        /// <param name="other">TeamMemberSettings Class object</param>
        /// <returns>true if the TeamMemberSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamMemberSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AllowCreateUpdateChannels == other.AllowCreateUpdateChannels &&
                this.AllowDeleteChannels == other.AllowDeleteChannels &&
                this.AllowAddRemoveApps == other.AllowAddRemoveApps &&
                this.AllowCreateUpdateRemoveTabs == other.AllowCreateUpdateRemoveTabs &&
                this.AllowCreateUpdateRemoveConnectors == other.AllowCreateUpdateRemoveConnectors &&
                this.AllowCreatePrivateChannels == other.AllowCreatePrivateChannels
                );
        }

        #endregion
    }
}
