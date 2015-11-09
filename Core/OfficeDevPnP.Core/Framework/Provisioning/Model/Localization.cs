using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Represents a Field XML Markup that is used to define information about a field
    /// </summary>
    public partial class Localization : IEquatable<Localization>
    {
        #region Private Members

        private Guid _id = Guid.Empty;
        #endregion

        #region Public Properties

        /// <summary>
        /// Gets ot sets the field ID if the Localization is for a Field
        /// </summary>
        public Guid Id
        {
            get { return this._id; }
            set { this._id = value; }
        }

        /// <summary>
        /// Gets or sets the CultureName
        /// </summary>
        public string CultureName { get; private set; }

        /// <summary>
        /// Gets or sets translation for Title
        /// </summary>
        public string TitleResource { get; set; }

        /// <summary>
        /// Gets or sets translation for Description
        /// </summary>
        public string DescriptionResource { get; set; }
        #endregion

        #region Constructors

        public Localization()
        {
        }

        public Localization(string cultureName)
        {
            this.CultureName = cultureName;
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (string.Format("{0}|{1}|{2}|{3}|",
                (this.Id != null ? this.Id.GetHashCode() : 0),
                this.CultureName.GetHashCode(),
                this.TitleResource.GetHashCode(),
                this.DescriptionResource.GetHashCode()
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Localization))
            {
                return (false);
            }
            return (Equals((Localization)obj));
        }

        public bool Equals(Localization other)
        {
            return (this.Id == other.Id &&
                this.CultureName == other.CultureName &&
                this.TitleResource == other.TitleResource &&
                this.DescriptionResource == other.DescriptionResource);
        }

        #endregion
    }
}
