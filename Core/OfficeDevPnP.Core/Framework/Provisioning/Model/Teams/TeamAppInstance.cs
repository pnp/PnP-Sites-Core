using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    public partial class TeamAppInstance : BaseModel, IEquatable<TeamAppInstance>
    {
        #region Public Members

        /// <summary>
        /// Defines the unique ID of the App to install or update on the Team
        /// </summary>
        public String AppId { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                AppId?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamAppInstance class
        /// </summary>
        /// <param name="obj">Object that represents TeamAppInstance</param>
        /// <returns>Checks whether object is TeamAppInstance class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamAppInstance))
            {
                return (false);
            }
            return (Equals((TeamAppInstance)obj));
        }

        /// <summary>
        /// Compares TeamAppInstance object based on AppId
        /// </summary>
        /// <param name="other">TeamAppInstance Class object</param>
        /// <returns>true if the TeamAppInstance object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamAppInstance other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AppId == other.AppId);
        }

        #endregion
    }
}
