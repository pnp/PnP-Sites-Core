using System;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Defines a Unified Group
    /// </summary>
    public class UnifiedGroupEntity
    {
        /// <summary>
        /// Unified group id
        /// </summary>
        public String GroupId { get; set; }
        /// <summary>
        /// Unified group display name
        /// </summary>
        public String DisplayName { get; set; }
        /// <summary>
        /// Unified group description 
        /// </summary>
        public String Description { get; set; }
        /// <summary>
        /// Unified group mail
        /// </summary>
        public String Mail { get; set; }
        /// <summary>
        /// Unified group nick name
        /// </summary>
        public String MailNickname { get; set; }
        /// <summary>
        /// Url of site to configure unified group
        /// </summary>
        public String SiteUrl { get; set; }
        /// <summary>
        /// Classification of the Office 365 group
        /// </summary>
        public String Classification { get; set; }
        /// <summary>
        /// Visibility of the Office 365 group
        /// </summary>
        public String Visibility { get; set; }
        /// <summary>
        /// Indication if the Office 365 Group has a Microsoft Team provisioned for it
        /// </summary>
        public bool? HasTeam { get; set; }
    }
}
