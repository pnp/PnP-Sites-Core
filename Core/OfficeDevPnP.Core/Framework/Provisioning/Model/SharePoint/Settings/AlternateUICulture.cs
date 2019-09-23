using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{ 
    /// <summary>
    /// Defines an AlternateUICulture item for a web site settings
    /// </summary>
    public partial class AlternateUICulture : BaseModel, IEquatable<AlternateUICulture>
    {
        #region Properties

        /// <summary>
        /// The Locale ID of a AlternateUICulture
        /// </summary>
        public Int32 LCID { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for AlternateUICulture class
        /// </summary>
        public AlternateUICulture() { }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                (this.LCID.GetHashCode())
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with AlternateUICulture
        /// </summary>
        /// <param name="obj">Object that represents AlternateUICulture</param>
        /// <returns>true if the current object is equal to the AlternateUICulture</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is AlternateUICulture))
            {
                return (false);
            }
            return (Equals((AlternateUICulture)obj));
        }

        /// <summary>
        /// Compares AlternateUICulture object based on LCID
        /// </summary>
        /// <param name="other">AlternateUICulture object</param>
        /// <returns>true if the AlternateUICulture object is equal to the current object; otherwise, false.</returns>
        public bool Equals(AlternateUICulture other)
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
