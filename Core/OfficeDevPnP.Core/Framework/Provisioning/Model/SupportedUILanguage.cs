
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class SupportedUILanguage : IEquatable<SupportedUILanguage>
    {
        public Int32 LCID { get; set; }

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}",
                this.LCID
                ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is SupportedUILanguage))
            {
                return (false);
            }
            return (Equals((SupportedUILanguage)obj));
        }

        public bool Equals(SupportedUILanguage other)
        {
            return (this.LCID == other.LCID);
        }

        #endregion
    }
}
