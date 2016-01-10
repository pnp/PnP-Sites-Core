using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that represents an Feature.
    /// </summary>
    public partial class Feature : BaseModel, IEquatable<Feature>
    {
        #region Private Members

        private Guid _id;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the feature Id
        /// </summary>
        public Guid Id { get { return _id; } set { _id = value; } }

        /// <summary>
        /// Gets or sets if the feature should be deactivated
        /// </summary>
        public bool Deactivate { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                this.Deactivate.GetHashCode(),
                (this.Id != null ? this.Id.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Feature))
            {
                return (false);
            }
            return (Equals((Feature)obj));
        }

        public bool Equals(Feature other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Deactivate == other.Deactivate &&
                this.Id == other.Id);
        }

        #endregion
    }
}
