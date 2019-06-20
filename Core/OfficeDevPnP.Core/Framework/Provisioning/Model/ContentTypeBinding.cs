using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object for Content Type Binding in the Provisioning Template 
    /// </summary>
    public partial class ContentTypeBinding : BaseModel, IEquatable<ContentTypeBinding>
    {
        #region Private Members

        private string _contentTypeId;

        #endregion

        #region Properties
        /// <summary>
        /// Gets or Sets the Content Type ID 
        /// </summary>
        public string ContentTypeId { get { return _contentTypeId; } set { _contentTypeId = value; } }

        /// <summary>
        /// Gets or Sets if the Content Type should be the default Content Type in the library
        /// </summary>
        public bool Default { get; set; }

        /// <summary>
        /// Declares if the Content Type should be Removed from the list or library
        /// </summary>
        public bool Remove { get; set; } = false;

        /// <summary>
        /// Declares if the Content Type should be Hidden from New button of the list or library, optional attribute.
        /// </summary>
        public bool Hidden { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                (this.ContentTypeId != null ? this.ContentTypeId.GetHashCode() : 0),
                this.Default.GetHashCode(),
                this.Remove.GetHashCode(),
                this.Hidden.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ContentTypeBinding
        /// </summary>
        /// <param name="obj">Object that represents ContentTypeBinding</param>
        /// <returns>true if the current object is equal to the ContentTypeBinding</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ContentTypeBinding))
            {
                return (false);
            }
            return (Equals((ContentTypeBinding)obj));
        }

        /// <summary>
        /// Compares ContentTypeBinding object based on ContentTypeId, Default and Remove properties.
        /// </summary>
        /// <param name="other">ContentTypeBinding object</param>
        /// <returns>true if the ContentTypeBinding object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ContentTypeBinding other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ContentTypeId == other.ContentTypeId &&
                this.Default == other.Default &&
                this.Remove == other.Remove &&
                this.Hidden == other.Hidden
                );
        }

        #endregion
    }
}
