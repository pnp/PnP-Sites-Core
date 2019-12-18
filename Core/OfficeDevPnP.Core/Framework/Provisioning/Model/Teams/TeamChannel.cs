using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines a Channel for a Team
    /// </summary>
    public partial class TeamChannel : BaseModel, IEquatable<TeamChannel>
    {
        #region Public Members

        /// <summary>
        /// Defines a collection of Tabs for a Channel in a Team
        /// </summary>
        public TeamTabCollection Tabs { get; private set; }

        /// <summary>
        /// Defines a collection of Resources for Tabs in a Team Channel
        /// </summary>
        public TeamTabResourceCollection TabResources { get; private set; }

        /// <summary>
        /// Defines a collection of Messages for a Team Channe
        /// </summary>
        public TeamChannelMessageCollection Messages { get; private set; }

        /// <summary>
        /// Defines the Display Name of the Channel
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// Defines the Description of the Channel
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// Defines whether the Channel is Favorite by default for all members of the Team
        /// </summary>
        public Boolean? IsFavoriteByDefault { get; set; }

        /// <summary>
        /// Declares the ID for the Channel
        /// </summary>
        public String ID { get; set; }        

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for TeamChannel
        /// </summary>
        public TeamChannel()
        {
            this.Tabs = new TeamTabCollection(this.ParentTemplate);
            this.TabResources = new TeamTabResourceCollection(this.ParentTemplate);
            this.Messages = new TeamChannelMessageCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|",
                Tabs.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                TabResources.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                Messages.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                DisplayName?.GetHashCode() ?? 0,
                Description?.GetHashCode() ?? 0,
                IsFavoriteByDefault.GetHashCode(),
                ID?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamChannel class
        /// </summary>
        /// <param name="obj">Object that represents TeamChannel</param>
        /// <returns>Checks whether object is TeamChannel class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamChannel))
            {
                return (false);
            }
            return (Equals((TeamChannel)obj));
        }

        /// <summary>
        /// Compares TeamChannel object based on Tabs, TabResources, Messages, DisplayName, Description, and IsFavoriteByDefault
        /// </summary>
        /// <param name="other">TeamChannel Class object</param>
        /// <returns>true if the TeamChannel object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamChannel other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Tabs.DeepEquals(other.Tabs) &&
                this.TabResources.DeepEquals(other.TabResources) &&
                this.Messages.DeepEquals(other.Messages) &&
                this.DisplayName == other.DisplayName &&
                this.Description == other.Description &&
                this.IsFavoriteByDefault == other.IsFavoriteByDefault &&
                this.ID == other.ID
                );
        }

        #endregion
    }
}
