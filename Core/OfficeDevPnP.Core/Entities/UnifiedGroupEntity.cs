using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
