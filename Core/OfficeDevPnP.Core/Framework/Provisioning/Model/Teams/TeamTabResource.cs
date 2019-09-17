using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines a Resource for a Tab in a Channel of a Team
    /// </summary>
    public partial class TeamTabResource : BaseModel, IEquatable<TeamTabResource>
    {
        #region Public Members

        /// <summary>
        /// Defines the Configuration for the Tab Resource
        /// </summary>
        public String TabResourceSettings { get; set; }

        /// <summary>
        /// Defines the Type of Resource for the Tab
        /// </summary>
        public TabResourceType Type { get; set; }

        /// <summary>
        /// Defines the ID of the target Tab for the Resource
        /// </summary>
        public String TargetTabId { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                TabResourceSettings?.GetHashCode() ?? 0,
                Type.GetHashCode(),
                TargetTabId?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamTabResource class
        /// </summary>
        /// <param name="obj">Object that represents TeamTabResource</param>
        /// <returns>Checks whether object is TeamTabResource class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamTabResource))
            {
                return (false);
            }
            return (Equals((TeamTabResource)obj));
        }

        /// <summary>
        /// Compares TeamTabResource object based on TabResourceSettings, Type, and TargetTabId
        /// </summary>
        /// <param name="other">TeamTabResource Class object</param>
        /// <returns>true if the TeamTabResource object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamTabResource other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.TabResourceSettings == other.TabResourceSettings &&
                this.Type == other.Type &&
                this.TargetTabId == other.TargetTabId
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the Types of Resources for the Tab
    /// </summary>
    public enum TabResourceType
    {
        /// <summary>
        /// Defines a Generic resource type
        /// </summary>
        Generic,
        /// <summary>
        /// Defines a Notebook resource type
        /// </summary>
        Notebook,
        /// <summary>
        /// Defines a Planner resource type
        /// </summary>
        Planner,
        /// <summary>
        /// Defines a Schedule resource type
        /// </summary>
        Schedule,
    }
}
