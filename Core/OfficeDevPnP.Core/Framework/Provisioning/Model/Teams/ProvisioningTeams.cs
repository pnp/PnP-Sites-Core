using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines a complex type declaring settings for provisioning Microsoft Teams objects
    /// </summary>
    public partial class ProvisioningTeams : BaseModel, IEquatable<ProvisioningTeams>
    {
        #region Public Members

        /// <summary>
        /// A collection of Teams to provision starting from a Template
        /// </summary>
        public TeamTemplateCollection TeamTemplates { get; private set; }

        /// <summary>
        /// A collection of Teams to provision/update
        /// </summary>
        public TeamCollection Teams { get; private set; }

        /// <summary>
        /// A collection of Team Apps to provision
        /// </summary>
        public TeamAppCollection Apps { get; private set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for ProvisioningTeams
        /// </summary>
        public ProvisioningTeams()
        {
            this.TeamTemplates = new TeamTemplateCollection(this.ParentTemplate);
            this.Teams = new TeamCollection(this.ParentTemplate);
            this.Apps = new TeamAppCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}",
                this.TeamTemplates.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Teams.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Apps.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ProvisioningTeams
        /// </summary>
        /// <param name="obj">Object that represents ProvisioningTeams</param>
        /// <returns>true if the current object is equal to the ProvisioningTeams</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningTeams))
            {
                return (false);
            }
            return (Equals((ProvisioningTeams)obj));
        }

        /// <summary>
        /// Compares ProvisioningTeams object based on TeamTemplates, Teams, and Apps
        /// </summary>
        /// <param name="other">ProvisioningTeams object</param>
        /// <returns>true if the ProvisioningTeams object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ProvisioningTeams other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.TeamTemplates.DeepEquals(other.TeamTemplates) &&
                this.Teams.DeepEquals(other.Teams) &&
                this.Apps.DeepEquals(other.Apps)
                );
        }

        #endregion
    }
}
