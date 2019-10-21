using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object used to define a Site Design Right Grant
    /// </summary>
    public partial class SiteDesignGrant : BaseModel, IEquatable<SiteDesignGrant>
    {
        #region Properties

        /// <summary>
        /// Gets or sets the Principal for the SiteDesignGrant
        /// </summary>
        public string Principal { get; set; }

        /// <summary>
        /// Gets or sets the Right for the SiteDesignGrant
        /// </summary>
        public SiteDesignRight Right { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                this.Principal?.GetHashCode() ?? 0,
                this.Right.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteDesignGrant
        /// </summary>
        /// <param name="obj">Object</param>
        /// <returns>true if the current object is equal to the SiteDesignGrant</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteDesignGrant))
            {
                return (false);
            }
            return (Equals((SiteDesignGrant)obj));
        }

        /// <summary>
        /// Compares SiteDesignGrant object based on Principal, and Right properties.
        /// </summary>
        /// <param name="other">SiteDesignGrant object</param>
        /// <returns>true if the SiteDesignGrant object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteDesignGrant other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Principal == other.Principal &&
                this.Right == other.Right
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the Right options for a SiteDesignGrant
    /// </summary>
    public enum SiteDesignRight
    {
        /// <summary>
        /// Right for None
        /// </summary>
        None,
        /// <summary>
        /// Right to View
        /// </summary>
        View,
    }
}
