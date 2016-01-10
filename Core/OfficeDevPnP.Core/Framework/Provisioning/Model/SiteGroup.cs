using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// The base type for a Site Group
    /// </summary>
    public partial class SiteGroup : BaseModel, IEquatable<SiteGroup>
    {
        #region Private Members

        private UserCollection _members;

        #endregion

        #region Constructors

        public SiteGroup()
        {
            this._members = new UserCollection(this.ParentTemplate);
        }

        public SiteGroup(IEnumerable<User> members):
            this()
        {
            this.Members.AddRange(members);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// The list of members of the Site Group
        /// </summary>
        public UserCollection Members
        {
            get { return this._members; }
            private set { this._members = value; }
        }

        /// <summary>
        /// The Title of the Site Group
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// The Description of the Site Group
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The Owner of the Site Group
        /// </summary>
        public String Owner { get; set; }

        /// <summary>
        /// Defines whether the members can edit membership of the Site Group
        /// </summary>
        public Boolean AllowMembersEditMembership { get; set; }

        /// <summary>
        /// Defines whether to allow requests to join or leave the Site Group
        /// </summary>
        public Boolean AllowRequestToJoinLeave { get; set; }

        /// <summary>
        /// Defines whether to auto-accept requests to join or leave the Site Group
        /// </summary>
        public Boolean AutoAcceptRequestToJoinLeave { get; set; }

        /// <summary>
        /// Defines whether to allow members only to view the membership of the Site Group
        /// </summary>
        public Boolean OnlyAllowMembersViewMembership { get; set; }

        /// <summary>
        /// Defines the email address used for membership requests to join or leave will be sent for the Site Group
        /// </summary>
        public String RequestToJoinLeaveEmailSetting { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|",
                this.AllowMembersEditMembership.GetHashCode(),
                this.AllowRequestToJoinLeave.GetHashCode(),
                this.AutoAcceptRequestToJoinLeave.GetHashCode(),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                this.Members.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.OnlyAllowMembersViewMembership.GetHashCode(),
                (this.Owner != null ? this.Owner.GetHashCode() : 0),
                (this.RequestToJoinLeaveEmailSetting != null ? this.RequestToJoinLeaveEmailSetting.GetHashCode() : 0),
                (this.Title != null ? this.Title.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is SiteGroup))
            {
                return (false);
            }
            return (Equals((SiteGroup)obj));
        }

        public bool Equals(SiteGroup other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.AllowMembersEditMembership == other.AllowMembersEditMembership &&
                this.AllowRequestToJoinLeave == other.AllowRequestToJoinLeave &&
                this.AutoAcceptRequestToJoinLeave == other.AutoAcceptRequestToJoinLeave &&
                this.Description == other.Description &&
                this.Members.DeepEquals(other.Members) &&
                this.OnlyAllowMembersViewMembership == other.OnlyAllowMembersViewMembership &&
                this.Owner == other.Owner &&
                this.RequestToJoinLeaveEmailSetting == other.RequestToJoinLeaveEmailSetting &&
                this.Title == other.Title
                );
        }

        #endregion
    }
}
