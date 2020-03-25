using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A welcome page field for a Document Set
    /// </summary>
    public partial class WelcomePageField : BaseModel, IEquatable<WelcomePageField>
    {
        #region Public Members

        /// <summary>
        /// The name of the welcome page field in a document set
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// The id of the welcome page field in a document set
        /// </summary>
        public Guid FieldId { get; set; }

        /// <summary>
        /// True to specify that the welcome page field should be removed from the document set. If False, it means it will be added to the document set.
        /// </summary>
        public bool Remove { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.FieldId != null ? this.FieldId.GetHashCode() : 0),
                this.Remove.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with WelcomePageField
        /// </summary>
        /// <param name="obj">Object that represents WelcomePageField</param>
        /// <returns>True if the current object is equal to the WelcomePageField</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is WelcomePageField))
            {
                return (false);
            }
            return (Equals((WelcomePageField)obj));
        }

        /// <summary>
        /// Compares WelcomePageField object based on Name, FieldId and Remove.
        /// </summary>
        /// <param name="other">WelcomePageField object</param>
        /// <returns>True if the WelcomePageField object is equal to the current object; otherwise, false.</returns>
        public bool Equals(WelcomePageField other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                    this.FieldId == other.FieldId &&
                    this.Remove == other.Remove
                );
        }

        #endregion
    }
}
