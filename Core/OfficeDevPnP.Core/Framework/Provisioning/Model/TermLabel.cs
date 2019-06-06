using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class TermLabel : BaseModel, IEquatable<TermLabel>
    {

        #region Public Members
        /// <summary>
        /// Gets or sets the Language for the term label
        /// </summary>
        public int Language { get; set; }
        /// <summary>
        /// Gets or sets the IsDefaultForLangauage flag for the term label
        /// </summary>
        public bool IsDefaultForLanguage { get; set; }
        /// <summary>
        /// Gets or sets the Value for the term label
        /// </summary>
        public string Value { get; set; }

        #endregion

        #region Constructors
        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                this.Language.GetHashCode(),
                (this.Value != null ? this.Value.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TermLabel
        /// </summary>
        /// <param name="obj">Object that represents TermLabel</param>
        /// <returns>true if the current object is equal to the TermLabel</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TermLabel))
            {
                return (false);
            }
            return (Equals((TermLabel)obj));
        }

        /// <summary>
        /// Compares TermLabel object based on Language and Value. 
        /// </summary>
        /// <param name="other">TermLabel object</param>
        /// <returns>true if the TermLabel object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TermLabel other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Language == other.Language &&
                this.Value == other.Value);
        }

        #endregion
    }
}
