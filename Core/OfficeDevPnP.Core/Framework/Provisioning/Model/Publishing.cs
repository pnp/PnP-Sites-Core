using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Publishing configuration to provision
    /// </summary>
    public partial class Publishing : BaseModel, IEquatable<Publishing>
    {
        #region Private Members

        private DesignPackage _designPackage = null;
        private AvailableWebTemplateCollection _availableWebTemplates;
        private PageLayoutCollection _pageLayouts;
        private ImageRenditionCollection _imageRenditions;

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for Publishing class
        /// </summary>
        public Publishing()
        {
            this._availableWebTemplates = new AvailableWebTemplateCollection(this.ParentTemplate);
            this._pageLayouts = new PageLayoutCollection(this.ParentTemplate);
            this._imageRenditions = new ImageRenditionCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for Publishing class
        /// </summary>
        /// <param name="autoCheckRequirements">AutoCheckRequirementsOption object</param>
        /// <param name="designPackage">Design Package for publishing</param>
        /// <param name="availableWebTemplates">Available WebTemplates for publishing</param>
        /// <param name="pageLayouts">PageLayouts for publishing</param>
        public Publishing(AutoCheckRequirementsOptions autoCheckRequirements, DesignPackage designPackage = null, IEnumerable<AvailableWebTemplate> availableWebTemplates = null, IEnumerable<PageLayout> pageLayouts = null) 
            : this()
        {
            this.AutoCheckRequirements = autoCheckRequirements;

            if (designPackage != null)
            {
                this.DesignPackage = designPackage;
            }
            this.AvailableWebTemplates.AddRange(availableWebTemplates);
            this.PageLayouts.AddRange(pageLayouts);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines a Design Package to import into the current Publishing site
        /// </summary>
        public DesignPackage DesignPackage
        {
            get { return this._designPackage; }
            set
            {
                if (this._designPackage != null)
                {
                    this._designPackage.ParentTemplate = null;
                }
                this._designPackage = value;
                if (this._designPackage != null)
                {
                    this._designPackage.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines the Available Web Templates for the current Publishing site
        /// </summary>
        public AvailableWebTemplateCollection AvailableWebTemplates
        {
            get { return this._availableWebTemplates; }
            private set { this._availableWebTemplates = value; }
        }

        /// <summary>
        /// Defines the Available Page Layouts for the current Publishing site
        /// </summary>
        public PageLayoutCollection PageLayouts
        {
            get { return this._pageLayouts; }
            private set { this._pageLayouts = value; }
        }

        /// <summary>
        /// Defines how an engine should behave if the requirements for provisioning publishing capabilities are not satisfied by the target site 
        /// </summary>
        public AutoCheckRequirementsOptions AutoCheckRequirements { get; set; }

        public ImageRenditionCollection ImageRenditions
        {
            get { return this._imageRenditions; }
            private set { this._imageRenditions = value; }
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|",
                this.AutoCheckRequirements.GetHashCode(),
                this.AvailableWebTemplates.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                (this.DesignPackage != null ? this.DesignPackage.GetHashCode() : 0),
                this.PageLayouts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.ImageRenditions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Publishing
        /// </summary>
        /// <param name="obj">Object that represents Publishing</param>
        /// <returns>true if the current object is equal to the Publishing</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Publishing))
            {
                return (false);
            }
            return (Equals((Publishing)obj));
        }

        /// <summary>
        /// Compares Publishing object based on AutoCheckRequirements, AvailableWebTemplates, DesignPackage and PageLayout properties.
        /// </summary>
        /// <param name="other">Publishing object</param>
        /// <returns>true if the Publishing object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Publishing other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.AutoCheckRequirements == other.AutoCheckRequirements &&
                this.AvailableWebTemplates.DeepEquals(other.AvailableWebTemplates) &&
                this.DesignPackage == other.DesignPackage &&
                this.PageLayouts.DeepEquals(other.PageLayouts) &&
                this.ImageRenditions.DeepEquals(other.ImageRenditions)
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines how an engine should behave if the requirements for provisioning publishing capabilities are not satisfied by the target site 
    /// </summary>
    public enum AutoCheckRequirementsOptions
    {
        /// <summary>
        /// Instructs the engine to make the target site compliant with the requirements
        /// </summary>
        MakeCompliant,
        /// <summary>
        /// Instructs the engine to skip the Publishing section if the target site is not compliant with the requirements
        /// </summary>
        SkipIfNotCompliant,
        /// <summary>
        /// Instructs the engine to throw an exception/failure if the target site is not compliant with the requirements
        /// </summary>
        FailIfNotCompliant,
    }
}
