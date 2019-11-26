using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a ClientSidePage
    /// </summary>
    public partial class ClientSidePage : BaseClientSidePage
    {
        #region Private Members

        private TranslatedClientSidePageCollection _translations;

        #endregion

        #region Public Members

        /// <summary>
        /// Gets or sets the translations
        /// </summary>
        public TranslatedClientSidePageCollection Translations
        {
            get { return _translations; }
            private set { _translations = value; }
        }

        /// <summary>
        /// Defines the Page Name of the Client Side Page, required attribute.
        /// </summary>
        public String PageName { get; set; }

        /// <summary>
        /// Defines whether to create translations of the current Client Side Page
        /// </summary>
        public bool CreateTranslations { get; set; }

        /// <summary>
        /// Defines the Locale ID of a Localization Language
        /// </summary>
        public int LCID { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for ClientSidePage class
        /// </summary>
        public ClientSidePage()
        {
            this._translations = new TranslatedClientSidePageCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                this.Translations.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.PageName?.GetHashCode() ?? 0,
                this.CreateTranslations.GetHashCode(),
                this.LCID.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares ClientSidePage object based on PageName, CreateTranslations, and Translations
        /// </summary>
        /// <param name="other">ClientSidePage Class object</param>
        /// <returns>true if the ClientSidePage object is equal to the current object; otherwise, false.</returns>
        protected override bool EqualsInherited(BaseClientSidePage other)
        {
            var otherTyped = other as ClientSidePage;

            if (otherTyped == null)
            {
                return (false);
            }

            return (this.Translations.DeepEquals(otherTyped.Translations) &&
                this.PageName == otherTyped.PageName &&
                this.CreateTranslations == otherTyped.CreateTranslations &&
                this.LCID == otherTyped.LCID
                );
        }

        #endregion
    }
}
