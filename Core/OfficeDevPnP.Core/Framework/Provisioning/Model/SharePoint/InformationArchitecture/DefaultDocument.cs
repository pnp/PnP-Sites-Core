using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A default document for a Document Set
    /// </summary>
    public partial class DefaultDocument : BaseModel, IEquatable<DefaultDocument>
    {
        #region Public Members

        /// <summary>
        /// The name (including the relative path) of the Default Document for a Document Set
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// The value of the ContentTypeID of the Default Document for the Document Set
        /// </summary>
        public String ContentTypeId { get; set; }

        /// <summary>
        /// The path of the file to upload as a Default Document for the Document Set
        /// </summary>
        public String FileSourcePath { get; set; }

        /// <summary>
        /// True to specify that the Default Document should be removed from the document set. If False, it means it will be added to the document set.
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
                (this.FileSourcePath != null ? this.FileSourcePath.GetHashCode() : 0),
                this.Remove.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with DefaultDocument
        /// </summary>
        /// <param name="obj">Object that represents DefaultDocument</param>
        /// <returns>true if the current object is equal to the DefaultDocument</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DefaultDocument))
            {
                return (false);
            }
            return (Equals((DefaultDocument)obj));
        }

        /// <summary>
        /// Compares DefaultDocument object based on Name, ContentTypeID, FileSourcePath and Remove.
        /// </summary>
        /// <param name="other">DefaultDocument object</param>
        /// <returns>true if the DefaultDocument object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DefaultDocument other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                    this.ContentTypeId == other.ContentTypeId &&
                    this.FileSourcePath == other.FileSourcePath &&
                    this.Remove == other.Remove
                );
        }

        #endregion
    }
}
