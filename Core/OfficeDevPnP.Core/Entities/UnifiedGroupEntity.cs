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
        public String GroupId { get; set; }

        public String DisplayName { get; set; }

        public String Description { get; set; }

        public String Mail { get; set; }

        public String MailNickname { get; set; }

        public String SiteUrl { get; set; }
    }
}
