using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// An allowed content type for a Document Set
    /// </summary>
    public partial class AllowedContentType : BaseModel, IEquatable<AllowedContentType>
    {
        #region Public Members

        /// <summary>
        /// The name of the allowed content type in a document set
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// The content type id of the allowed content type in a document set
        /// </summary>
        public String ContentTypeId { get; set; }

        /// <summary>
        /// True to specify that the allowed content type should be removed from the document set. If False, it means it will be added to the document set.
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
                (this.ContentTypeId != null ? this.ContentTypeId.GetHashCode() : 0),
                this.Remove.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with AllowedContentType
        /// </summary>
        /// <param name="obj">Object that represents AllowedContentType</param>
        /// <returns>True if the current object is equal to the AllowedContentType</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is AllowedContentType))
            {
                return (false);
            }
            return (Equals((AllowedContentType)obj));
        }

        /// <summary>
        /// Compares AllowedContentType object based on Name, ContentTypeID and Remove.
        /// </summary>
        /// <param name="other">AllowedContentType object</param>
        /// <returns>True if the AllowedContentType object is equal to the current object; otherwise, false.</returns>
        public bool Equals(AllowedContentType other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                    this.ContentTypeId == other.ContentTypeId &&
                    this.Remove == other.Remove
                );

        }

        #endregion
    }
}
