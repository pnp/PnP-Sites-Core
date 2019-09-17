using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.SharePoint.InformationArchitecture
{
    public partial class DataRowAttachment : BaseModel, IEquatable<DataRowAttachment>
    {
        #region Public members

        /// <summary>
        /// The Name of the File Attachment
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// The Src of the File
        /// </summary>
        public String Src { get; set; }

        /// <summary>
        /// Defines whether to overwrite an already existing file or not
        /// </summary>
        public Boolean Overwrite { get; set; }

        #endregion

        #region Constructors

        public DataRowAttachment() : base()
        {
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
                Name?.GetHashCode() ?? 0,
                Src?.GetHashCode() ?? 0,
                Overwrite.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with DataRowAttachment class
        /// </summary>
        /// <param name="obj">Object that represents DataRowAttachment</param>
        /// <returns>Checks whether object is DataRowAttachment class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DataRowAttachment))
            {
                return (false);
            }
            return (Equals((DataRowAttachment)obj));
        }

        /// <summary>
        /// Compares DataRowAttachment object based on Name, Src, and Overwrite
        /// </summary>
        /// <param name="other">User DataRowAttachment object</param>
        /// <returns>true if the DataRowAttachment object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DataRowAttachment other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                this.Src == other.Src &&
                this.Overwrite == other.Overwrite
                );
        }

        #endregion
    }
}
