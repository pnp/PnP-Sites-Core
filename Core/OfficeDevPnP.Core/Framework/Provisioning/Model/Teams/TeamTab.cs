using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines an TeamTab for automated provisiong of Microsoft Teams
    /// </summary>
    public class TeamTab : BaseModel, IEquatable<TeamTab>
    {
        #region Public Members

        /// <summary>
        /// Defines the Display Name of the Channel
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// App definition identifier of the tab
        /// </summary>
        public String TeamsAppId { get; set; }

        /// <summary>
        /// Defines the Configuration for the Tab
        /// </summary>
        public TeamTabConfiguration Configuration { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                DisplayName?.GetHashCode() ?? 0,
                TeamsAppId?.GetHashCode() ?? 0,
                Configuration?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamTab class
        /// </summary>
        /// <param name="obj">Object that represents TeamTab</param>
        /// <returns>Checks whether object is TeamTab class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamTab))
            {
                return (false);
            }
            return (Equals((TeamTab)obj));
        }

        /// <summary>
        /// Compares TeamTab object based on DisplayName, TeamsAppId, and Configuration
        /// </summary>
        /// <param name="other">TeamTab Class object</param>
        /// <returns>true if the TeamTab object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamTab other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.DisplayName == other.DisplayName &&
                this.TeamsAppId == other.TeamsAppId &&
                this.Configuration == other.Configuration
                );
        }

        #endregion
    }
}
