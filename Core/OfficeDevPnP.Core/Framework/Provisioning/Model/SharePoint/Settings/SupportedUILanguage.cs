
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
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                this.LCID.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SupportedUILanguage
        /// </summary>
        /// <param name="obj">Object that represents SupportedUILanguage</param>
        /// <returns>true if the current object is equal to the SupportedUILanguage</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SupportedUILanguage))
            {
                return (false);
            }
            return (Equals((SupportedUILanguage)obj));
        }

        /// <summary>
        /// Compares SupportedUILanguage object based on LCID.
        /// </summary>
        /// <param name="other">SupportedUILanguage object</param>
        /// <returns>true if the SupportedUILanguage object is equal to the current object; otherwise, false.</returns>
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
