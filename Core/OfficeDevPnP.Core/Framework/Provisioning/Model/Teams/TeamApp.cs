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
    public partial class TeamApp : BaseModel, IEquatable<TeamApp>
    {
        #region Public Members

        /// <summary>
        /// Unique ID - from PnP perspective - for the App, defined for further reference in the Provisioning Template
        /// </summary>
        public String AppId { get; set; }

        /// <summary>
        /// The URL or path for the Teams App package
        /// </summary>
        public String PackageUrl { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                AppId?.GetHashCode() ?? 0,
                PackageUrl?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamApp class
        /// </summary>
        /// <param name="obj">Object that represents TeamApp</param>
        /// <returns>Checks whether object is TeamApp class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamApp))
            {
                return (false);
            }
            return (Equals((TeamApp)obj));
        }

        /// <summary>
        /// Compares TeamApp object based on AppId, and PackageUrl
        /// </summary>
        /// <param name="other">TeamApp Class object</param>
        /// <returns>true if the TeamApp object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamApp other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AppId == other.AppId &&
                this.PackageUrl == other.PackageUrl
                );
        }

        #endregion
    }
}
