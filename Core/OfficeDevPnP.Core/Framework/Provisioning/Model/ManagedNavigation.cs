using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Managed Navigation settings of a site
    /// </summary>
    public partial class ManagedNavigation : BaseModel, IEquatable<ManagedNavigation>
    {
        #region Public Members

        /// <summary>
        /// Defines the TermStore ID for the Managed Navigation
        /// </summary>
        public String TermStoreId { get; set; }

        /// <summary>
        /// Defines the TermSet ID for the Managed Navigation
        /// </summary>
        public String TermSetId { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                (this.TermStoreId != null ? this.TermStoreId.GetHashCode() : 0),
                (this.TermSetId != null ? this.TermSetId.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ManagedNavigation
        /// </summary>
        /// <param name="obj">Object the represents ManagedNavigation</param>
        /// <returns>true if the current object is equal to the ManagedNavigation</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ManagedNavigation))
            {
                return (false);
            }
            return (Equals((ManagedNavigation)obj));
        }

        /// <summary>
        /// Compares ManagedNavigation object based on TermStoreId and TermSetId
        /// </summary>
        /// <param name="other">ManagedNavigation object</param>
        /// <returns>Returns true if it matches with id of termstore and termset</returns>
        public bool Equals(ManagedNavigation other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.TermStoreId == other.TermStoreId &&
                this.TermSetId == other.TermSetId
                );
        }

        #endregion
    }
}
