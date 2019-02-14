using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a sequence of activities to execute with the engine
    /// </summary>
    /// <remarks>
    /// Each Provisioning file is split into a set of Sequence elements.
    /// The Sequence element groups the artefacts to be provisioned into groups.
    /// The Sequences must be evaluated by the provisioning engine in the order in which they appear.
    /// </remarks>
    public partial class ProvisioningSequence : BaseHierarchyModel, IEquatable<ProvisioningSequence>
    {
        #region Private Members

        private ProvisioningTermStore _termStore;

        #endregion

        #region Constructors

        public ProvisioningSequence()
        {
            this.SiteCollections = new SiteCollectionCollection(this.ParentHierarchy);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// A unique identifier of the Sequence, required attribute.
        /// </summary>
        public String ID { get; set; }

        public SiteCollectionCollection SiteCollections { get; private set; }

        /// <summary>
        /// Defines the TermStore to provision, if any
        /// </summary>
        public ProvisioningTermStore TermStore
        {
            get { return this._termStore; }
            set { this._termStore = value; }
        }

        public override string ToString()
        {
            return ID;
        }
        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.ID?.GetHashCode() ?? 0,
                this.TermStore?.GetHashCode() ?? 0,
                this.SiteCollections.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ProvisioningSequence
        /// </summary>
        /// <param name="obj">Object that represents ProvisioningSequence</param>
        /// <returns>true if the current object is equal to the ProvisioningSequence</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningSequence))
            {
                return (false);
            }
            return (Equals((ProvisioningSequence)obj));
        }

        /// <summary>
        /// Compares ProvisioningSequence object based on its properties
        /// </summary>
        /// <param name="other">ProvisioningSequence object</param>
        /// <returns>true if the ProvisioningSequence object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ProvisioningSequence other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ID == other.ID &&
                this.TermStore == other.TermStore &&
                this.SiteCollections.DeepEquals(other.SiteCollections)
                );
        }

        #endregion
    }
}
