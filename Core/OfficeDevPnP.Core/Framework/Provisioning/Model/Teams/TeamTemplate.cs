using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines a Team Template for automated provisiong of Microsoft Teams
    /// </summary>
    public partial class TeamTemplate : BaseTeam, IEquatable<TeamTemplate>
    {
        #region Public Members

        /// <summary>
        /// The JSON content of the Team Template
        /// </summary>
        public String JsonTemplate { get; set; }

        #endregion

        #region Constructors
        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                base.GetHashCode(),
                JsonTemplate?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamTemplate class
        /// </summary>
        /// <param name="obj">Object that represents TeamTemplate</param>
        /// <returns>Checks whether object is TeamTemplate class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamTemplate))
            {
                return (false);
            }
            return (Equals((TeamTemplate)obj));
        }

        /// <summary>
        /// Compares TeamTemplate object based on JsonTemplate
        /// </summary>
        /// <param name="other">TeamTemplate Class object</param>
        /// <returns>true if the TeamTemplate object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamTemplate other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.JsonTemplate == other.JsonTemplate &&
                base.Equals(other)
                );
        }

        #endregion
    }
}
