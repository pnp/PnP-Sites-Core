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

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                this.LanguageCode.GetHashCode(),
                (this.TemplateName != null ? this.TemplateName.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is AvailableWebTemplate))
            {
                return (false);
            }
            return (Equals((AvailableWebTemplate)obj));
        }

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
