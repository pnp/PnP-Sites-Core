using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a CdnOrigin 
    /// </summary>
    public partial class CdnOrigin : BaseModel, IEquatable<CdnOrigin>
    {
        #region Public Members

        /// <summary>
        /// Defines the URL for the CDN Origin, required attribute.
        /// </summary>
        public String Url { get; set; }

        /// <summary>
        /// Defines the action to perform with the CDN Origin for the CDN, required attribute.
        /// </summary>
        public OriginAction Action { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                this.Url?.GetHashCode() ?? 0,
                this.Action.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with CdnOrigin class
        /// </summary>
        /// <param name="obj">Object that represents CdnOrigin</param>
        /// <returns>Checks whether object is CdnOrigin class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is CdnOrigin))
            {
                return (false);
            }
            return (Equals((CdnOrigin)obj));
        }

        /// <summary>
        /// Compares CdnOrigin object based on Url, Action
        /// </summary>
        /// <param name="other">CdnOrigin Class object</param>
        /// <returns>true if the CdnOrigin object is equal to the current object; otherwise, false.</returns>
        public bool Equals(CdnOrigin other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Url == other.Url &&
                this.Action == other.Action
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the action to perform with the CDN Origin for the CDN
    /// </summary>
    public enum OriginAction
    {
        /// <summary>
        /// Declares to add a CDN Origin for the CDN.
        /// </summary>
        Add,
        /// <summary>
        /// Declares to remove a CDN Origin for the CDN.
        /// </summary>
        Remove,
    }
}
