using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a SharingSettings object
    /// </summary>
    public partial class SharingSettings : BaseModel, IEquatable<SharingSettings>
    {
        #region Public members

        /// <summary>
        /// Configures the sharing capability for the tenant
        /// </summary>
        public SharingCapability SharingCapability { get; set; }

        /// <summary>
        /// Number of days before expiration of anonymous sharing links
        /// </summary>
        public int RequireAnonymousLinksExpireInDays { get; set; }

        /// <summary>
        /// Defines the permissions for anonymous links for files
        /// </summary>
        public AnonymousLinkType FileAnonymousLinkType { get; set; }

        /// <summary>
        /// Defines the permissions for anonymous links for folders
        /// </summary>
        public AnonymousLinkType FolderAnonymousLinkType { get; set; }

        /// <summary>
        /// Defines the default type of a sharing link
        /// </summary>
        public SharingLinkType DefaultSharingLinkType { get; set; }

        /// <summary>
        /// Defines whether external users are allowed to reshare the content
        /// </summary>
        public bool PreventExternalUsersFromResharing { get; set; }

        /// <summary>
        /// Defines whether invited external users need to use the same account used as the target for the invite
        /// </summary>
        public bool RequireAcceptingAccountMatchInvitedAccount { get; set; }

        /// <summary>
        /// Defines domains restrictions for sharing
        /// </summary>
        public SharingDomainRestrictionMode SharingDomainRestrictionMode { get; set; }

        /// <summary>
        /// Defines a comma separated list of allowed domains for sharing. It is considered if and only if SharingDomainRestrictionMode=AllowList.
        /// </summary>
        public List<String> AllowedDomainList { get; internal set; } = new List<string>();

        /// <summary>
        /// Defines a comma separated list of blocked domains for sharing. It is considered if and only if SharingDomainRestrictionMode=BlockList.
        /// </summary>
        public List<String> BlockedDomainList { get; internal set; } = new List<string>();

        #endregion

        #region Constructors

        public SharingSettings() : base()
        {
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|",
                this.SharingCapability.GetHashCode(),
                this.RequireAnonymousLinksExpireInDays.GetHashCode(),
                this.FileAnonymousLinkType.GetHashCode(),
                this.FolderAnonymousLinkType.GetHashCode(),
                this.DefaultSharingLinkType.GetHashCode(),
                this.PreventExternalUsersFromResharing.GetHashCode(),
                this.RequireAcceptingAccountMatchInvitedAccount.GetHashCode(),
                this.SharingDomainRestrictionMode.GetHashCode(),
                this.AllowedDomainList.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.BlockedDomainList.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode()) ;
        }

        /// <summary>
        /// Compares object with SharingSettings class
        /// </summary>
        /// <param name="obj">Object that represents SharingSettings</param>
        /// <returns>Checks whether object is SharingSettings class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SharingSettings))
            {
                return (false);
            }
            return (Equals((SharingSettings)obj));
        }

        /// <summary>
        /// Compares SharingSettings object based on SharingCapability, RequireAnonymousLinksExpireInDays,
        /// FileAnonymousLinkType, FolderAnonymousLinkType, DefaultSharingLinkType, PreventExternalUsersFromResharing
        /// RequireAcceptingAccountMatchInvitedAccount, SharingDomainRestrictionMode, AllowedDomainList,
        /// and BlockedDomainList
        /// </summary>
        /// <param name="other">SharingSettings Class object</param>
        /// <returns>true if the SharingSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SharingSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.SharingCapability == other.SharingCapability &&
                this.RequireAnonymousLinksExpireInDays == other.RequireAnonymousLinksExpireInDays &&
                this.FileAnonymousLinkType == other.FileAnonymousLinkType &&
                this.FolderAnonymousLinkType == other.FolderAnonymousLinkType &&
                this.DefaultSharingLinkType == other.DefaultSharingLinkType &&
                this.PreventExternalUsersFromResharing == other.PreventExternalUsersFromResharing &&
                this.RequireAcceptingAccountMatchInvitedAccount == other.RequireAcceptingAccountMatchInvitedAccount &&
                this.SharingDomainRestrictionMode == other.SharingDomainRestrictionMode &&
                this.AllowedDomainList.DeepEquals(other.AllowedDomainList) &&
                this.BlockedDomainList.DeepEquals(other.BlockedDomainList)
                );
        }

        #endregion
    }

    public enum SharingCapability
    {
        /// <summary>
        /// Don't allow sharing outside your organization
        /// </summary>
        Disabled,
        /// <summary>
        /// Allow external users who accept sharing invitations and sign in as authenticated users
        /// </summary>
        ExternalUserSharingOnly,
        /// <summary>
        /// Allow sharing with all external users, and by using anonymous access links
        /// </summary>
        ExternalUserAndGuestSharing,
        /// <summary>
        /// Allow sharing only with the external users that already exist in your organization's directory
        /// </summary>
        ExistingExternalUserSharingOnly,
    }

    public enum AnonymousLinkType
    {
        /// <summary>
        /// No anonymous link type
        /// </summary>
        None,
        /// <summary>
        /// View permissions for anonymous links
        /// </summary>
        View,
        /// <summary>
        /// Edit permissions for anonymous links
        /// </summary>
        Edit,
    }

    public enum SharingLinkType
    {
        /// <summary>
        /// None
        /// </summary>
        None,
        /// <summary>
        /// Direct - only people who have permission
        /// </summary>
        Direct,
        /// <summary>
        /// Internal -  only people in the organization
        /// </summary>
        Internal,
        /// <summary>
        /// AnonymousAccess - anyone with the link
        /// </summary>
        AnonymousAccess,
    }

    public enum SharingDomainRestrictionMode
    {
        /// <summary>
        /// There are no domain restrictions in place
        /// </summary>
        None,
        /// <summary>
        /// The domain restrictions are based on an allowed domains list
        /// </summary>
        AllowList,
        /// <summary>
        /// The domain restrictions are based on an blocked domains list
        /// </summary>
        BlockList,
    }
}
