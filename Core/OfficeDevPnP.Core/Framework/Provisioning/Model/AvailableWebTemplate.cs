using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines an available Web Template for the current Publishing site
    /// </summary>
    public partial class AvailableWebTemplate : BaseModel, IEquatable<AvailableWebTemplate>
    {
        #region Public Members

        /// <summary>
        /// The Language Code for the Web Template
        /// </summary>
        public Int32 LanguageCode { get; set; }

        /// <summary>
        /// The Name of the Web Template
        /// </summary>
        public String TemplateName { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code.
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                this.LanguageCode.GetHashCode(),
                (this.TemplateName != null ? this.TemplateName.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with AvailableWebTemplate
        /// </summary>
        /// <param name="obj">Object that represents AvailableWebTemplate</param>
        /// <returns>true if the current object is equal to the AvailableWebTemplate</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is AvailableWebTemplate))
            {
                return (false);
            }
            return (Equals((AvailableWebTemplate)obj));
        }

        /// <summary>
        /// Compares AvailableWebTemplate object based on LanguageCode and TemplateName
        /// </summary>
        /// <param name="other">AvailableWebTemplate object</param>
        /// <returns>true if the AvailableWebTemplate object is equal to the current object; otherwise, false.</returns>
        public bool Equals(AvailableWebTemplate other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.LanguageCode == other.LanguageCode &&
                this.TemplateName == other.TemplateName
                );
        }

        #endregion
    }
}
