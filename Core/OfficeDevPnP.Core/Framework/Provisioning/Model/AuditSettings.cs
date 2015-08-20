using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class AuditSettings : IEquatable<AuditSettings>
    {
        public Microsoft.SharePoint.Client.AuditMaskType AuditFlag { get; set; }
        public Int32 AuditLogTrimmingRetention { get; set; }
        public Boolean TrimAuditLog { get; set; }

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}",
                this.AuditFlag,
                this.AuditLogTrimmingRetention,
                this.TrimAuditLog
                ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is AuditSettings))
            {
                return (false);
            }
            return (Equals((AuditSettings)obj));
        }

        public bool Equals(AuditSettings other)
        {
            return (this.AuditFlag == other.AuditFlag  &&
                this.AuditLogTrimmingRetention == other.AuditLogTrimmingRetention &&
                this.TrimAuditLog == other.TrimAuditLog
                );
        }

        #endregion
    }
}
