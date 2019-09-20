using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Term Store to provision through a Sequence, optional element.
    /// </summary>
    public partial class ProvisioningTermStore : IEquatable<ProvisioningTermStore>
    {
        #region Private Members

        private TermGroupCollection _termGroups;

        #endregion

        #region Constructor

        public ProvisioningTermStore()
        {
            // We don't have a parent Provisioning Template at this level
            this._termGroups = new TermGroupCollection(null);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Gets a collection of termgroups to deploy
        /// </summary>
        public TermGroupCollection TermGroups
        {
            get { return this._termGroups; }
            private set { this._termGroups = value; }
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                this.TermGroups.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TermStore
        /// </summary>
        /// <param name="obj">Object that represents TermStore</param>
        /// <returns>true if the current object is equal to the TermStore</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningTermStore))
            {
                return (false);
            }
            return (Equals((ProvisioningTermStore)obj));
        }

        /// <summary>
        /// Compares TermStore object based on its properties
        /// </summary>
        /// <param name="other">TermStore object</param>
        /// <returns>true if the TermStore object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ProvisioningTermStore other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.TermGroups.DeepEquals(other.TermGroups)
                );
        }

        #endregion
    }
}
