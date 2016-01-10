
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a single Supported UI Language for a site
    /// </summary>
    public partial class SupportedUILanguage : BaseModel, IEquatable<SupportedUILanguage>
    {
        /// <summary>
        /// The Locale ID of a Supported UI Language
        /// </summary>
        public Int32 LCID { get; set; }

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                this.LCID.GetHashCode()
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
            if (other == null)
            {
                return (false);
            }

            return (this.LCID == other.LCID);
        }

        #endregion
    }
}
