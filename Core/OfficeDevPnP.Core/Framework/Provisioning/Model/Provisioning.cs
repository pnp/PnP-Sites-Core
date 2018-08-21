using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the root node of the Provisioning Domain Model
    /// </summary>
    public partial class Provisioning
    {
        #region Private Fields

        private Dictionary<string, string> _parameters = new Dictionary<string, string>();
        private LocalizationCollection _localizations;
        private ProvisioningTenant _tenant;

        #endregion

        #region Constructors

        public Provisioning()
        {
            this.Templates = new ProvisioningTemplateCollection(this);
            this.Sequences = new ProvisioningSequenceCollection(this);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Any parameters that can be used throughout the template
        /// </summary>
        public Dictionary<string, string> Parameters
        {
            get { return _parameters; }
            private set { _parameters = value; }
        }

        /// <summary>
        /// Gets or sets the Localizations
        /// </summary>
        public LocalizationCollection Localizations
        {
            get { return this._localizations; }
            private set { this._localizations = value; }
        }

        /// <summary>
        /// The Tenant-wide settings for the template
        /// </summary>
        public ProvisioningTenant Tenant
        {
            get { return this._tenant; }
            set { this._tenant = value; }
        }

        /// <summary>
        /// Gets or sets the Provisioning File Version number
        /// </summary>
        public double Version { get; set; }

        /// <summary>
        /// Gets or sets the Provisioning File Author name
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Name of the tool generating this Provisioning File
        /// </summary>
        public string Generator { get; set; }

        /// <summary>
        /// The Description of the Provisioning File
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// The Display Name of the Provisioning File
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// The Image Preview Url of the Provisioning File
        /// </summary>
        public string ImagePreviewUrl { get; set; }

        /// <summary>
        /// A collection of Provisioning Template objects, if any
        /// </summary>
        public ProvisioningTemplateCollection Templates { get; private set; }

        /// <summary>
        /// A collection of Provisioning Sequence objects, if any
        /// </summary>
        public ProvisioningSequenceCollection Sequences { get; private set; }

        #endregion
    }
}
