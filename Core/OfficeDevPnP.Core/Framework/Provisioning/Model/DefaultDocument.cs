using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.ContentTypeId != null ? this.ContentTypeId.GetHashCode() : 0),
                (this.FileSourcePath != null ? this.FileSourcePath.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is DefaultDocument))
            {
                return (false);
            }
            return (Equals((DefaultDocument)obj));
        }

        public bool Equals(DefaultDocument other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                    this.ContentTypeId == other.ContentTypeId &&
                    this.FileSourcePath == other.FileSourcePath
                );

        }

        #endregion
    }
}
