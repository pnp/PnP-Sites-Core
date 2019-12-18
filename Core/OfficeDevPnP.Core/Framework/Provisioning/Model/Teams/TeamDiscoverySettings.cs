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
    public partial class TeamDiscoverySettings : BaseModel, IEquatable<TeamDiscoverySettings>
    {
        #region Public Members

        /// <summary>
        /// Defines whether the Team is visible via search and suggestions from the Teams client
        /// </summary>
        public Boolean ShowInTeamsSearchAndSuggestions { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}",
                ShowInTeamsSearchAndSuggestions.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamDiscoverySettings class
        /// </summary>
        /// <param name="obj">Object that represents TeamDiscoverySettings</param>
        /// <returns>Checks whether object is TeamDiscoverySettings class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamDiscoverySettings))
            {
                return (false);
            }
            return (Equals((TeamDiscoverySettings)obj));
        }

        /// <summary>
        /// Compares TeamDiscoverySettings object based on ShowInTeamsSearchAndSuggestions
        /// </summary>
        /// <param name="other">TeamDiscoverySettings Class object</param>
        /// <returns>true if the TeamDiscoverySettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamDiscoverySettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ShowInTeamsSearchAndSuggestions == other.ShowInTeamsSearchAndSuggestions
                );
        }

        #endregion
    }
}
