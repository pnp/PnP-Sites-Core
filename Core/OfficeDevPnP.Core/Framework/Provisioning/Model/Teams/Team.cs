using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines Team for automated provisiong/update of Microsoft Teams
    /// </summary>
    public partial class Team : BaseTeam, IEquatable<Team>
    {
        #region Public Members

        /// <summary>
        /// The Fun Settings for the Team
        /// </summary>
        public TeamFunSettings FunSettings { get; set; }

        /// <summary>
        /// The Guest Settings for the Team
        /// </summary>
        public TeamGuestSettings GuestSettings { get; set; }

        /// <summary>
        /// The Members Settings for the Team
        /// </summary>
        public TeamMemberSettings MemberSettings { get; set; }

        /// <summary>
        /// The Messaging Settings for the Team
        /// </summary>
        public TeamMessagingSettings MessagingSettings { get; set; }

        /// <summary>
        /// The Discovery Settings for the Team
        /// </summary>
        public TeamDiscoverySettings DiscoverySettings { get; set; }
        
        /// <summary>
        /// Defines the Security settings for the Team
        /// </summary>
        public TeamSecurity Security { get; set; }

        /// <summary>
        /// Defines the Channels for the Team
        /// </summary>
        public TeamChannelCollection Channels { get; private set; }

        /// <summary>
        /// Defines the Apps to install or update on the Team
        /// </summary>
        public TeamAppInstanceCollection Apps { get; private set; }

        public TeamSpecialization Specialization { get; set; }

        /// <summary>
        /// Declares the ID of the targt Group/Team to update, optional attribute. Cannot be used together with CloneFrom.
        /// </summary>
        public String GroupId { get; set; }

        /// <summary>
        /// Declares the ID of another Team to Clone the current Team from
        /// </summary>
        public String CloneFrom { get; set; }

        /// <summary>
        /// Declares whether the Team is archived or not
        /// </summary>
        public Boolean Archived { get; set; }

        /// <summary>
        /// Declares the nickname for the Team, optional attribute
        /// </summary>
        public String MailNickname { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for Team
        /// </summary>
        public Team()
        {
            this.Channels = new TeamChannelCollection(this.ParentTemplate);
            this.Apps = new TeamAppInstanceCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|",
                FunSettings?.GetHashCode() ?? 0,
                GuestSettings?.GetHashCode() ?? 0,
                MemberSettings?.GetHashCode() ?? 0,
                MessagingSettings?.GetHashCode() ?? 0,
                Security?.GetHashCode() ?? 0,
                Channels?.GetHashCode() ?? 0,
                Apps?.GetHashCode() ?? 0,
                Specialization.GetHashCode(),
                CloneFrom?.GetHashCode() ?? 0,
                Archived.GetHashCode(),
                GroupId?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Team class
        /// </summary>
        /// <param name="obj">Object that represents Team</param>
        /// <returns>Checks whether object is Team class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Team))
            {
                return (false);
            }
            return (Equals((Team)obj));
        }

        /// <summary>
        /// Compares Team object based on FunSettings, GuestSettings, MembersSettings, MessagingSettings, Security, Channels, Apps, Specialization, CloneFrom, Archived, and GroupId
        /// </summary>
        /// <param name="other">Team Class object</param>
        /// <returns>true if the Team object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Team other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.FunSettings == other.FunSettings &&
                this.GuestSettings == other.GuestSettings &&
                this.MemberSettings == other.MemberSettings &&
                this.MessagingSettings == other.MessagingSettings &&
                this.Security == other.Security &&
                this.Channels.DeepEquals(other.Channels) &&
                this.Apps.DeepEquals(other.Apps) &&
                this.Specialization == other.Specialization &&
                this.CloneFrom == other.CloneFrom &&
                this.Archived == other.Archived &&
                this.GroupId == other.GroupId
                );
        }

        #endregion
    }

    /// <summary>
    /// The Specialization for the Team
    /// </summary>
    public enum TeamSpecialization
    {
        /// <summary>
        /// Default type for a team which gives the standard team experience
        /// </summary>
        None,
        /// <summary>
        /// Team created by an education user. All teams created by education user are of type Edu.
        /// </summary>
        EducationStandard,
        /// <summary>
        /// Team experience optimized for a class. This enables segmentation of features across O365.
        /// </summary>
        EducationClass,
        /// <summary>
        /// Team experience optimized for a PLC. Learn more about PLC here.
        /// </summary>
        EducationProfessionalLearningCommunity,
        /// <summary>
        /// Team type for an optimized experience for staff in an organization, where a staff leader, like a principal, is the admin and teachers are members in a team that comes with a specialized notebook.
        /// </summary>
        EducationStaff,    
    }
}
