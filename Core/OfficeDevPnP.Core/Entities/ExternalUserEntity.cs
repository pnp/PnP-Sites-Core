using System;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Holds properties for external user entity
    /// </summary>
    public class ExternalUserEntity
    {
        /// <summary>
        /// External user accepted as with this value
        /// </summary>
        public string AcceptedAs { get; set; }
        /// <summary>
        /// External user display name
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// External user invited as with this value
        /// </summary>
        public string InvitedAs { get; set; }
        /// <summary>
        /// External user invited by this value
        /// </summary>
        public string InvitedBy { get; set; }
        /// <summary>
        /// Externla user unique id
        /// </summary>
        public string UniqueId { get; set; }
        /// <summary>
        /// Date Time of External user creation
        /// </summary>
        public DateTime WhenCreated { get; set; }

    }
}
