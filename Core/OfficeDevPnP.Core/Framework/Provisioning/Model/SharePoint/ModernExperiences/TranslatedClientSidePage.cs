using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Translated ClientSidePage
    /// </summary>
    public partial class TranslatedClientSidePage : BaseClientSidePage
    {
        #region Public Members

        /// <summary>
        /// Defines the Locale ID of a Localization Language
        /// </summary>
        public String LCID { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for TranslatedClientSidePage class
        /// </summary>
        public TranslatedClientSidePage()
        {
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|",
                this.LCID.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares TranslatedClientSidePage object based on LCID
        /// </summary>
        /// <param name="other">TranslatedClientSidePage Class object</param>
        /// <returns>true if the TranslatedClientSidePage object is equal to the current object; otherwise, false.</returns>
        protected override bool EqualsInherited(BaseClientSidePage other)
        {
            var otherTyped = other as TranslatedClientSidePage;

            if (otherTyped == null)
            {
                return (false);
            }

            return (this.LCID == otherTyped.LCID
                );
        }

        #endregion
    }
}
