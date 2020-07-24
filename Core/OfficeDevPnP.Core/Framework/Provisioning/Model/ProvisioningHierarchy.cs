using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the root node of the Provisioning Domain Model
    /// </summary>
    public partial class ProvisioningHierarchy
    {
        #region Constructors

        public ProvisioningHierarchy()
        {
            this.Templates = new ProvisioningTemplateCollection(this);
            this.Sequences = new ProvisioningSequenceCollection(this);
            this.Localizations = new LocalizationCollection(null);
            this.Tenant = new ProvisioningTenant();
            this.Teams = new Teams.ProvisioningTeams();
            this.AzureActiveDirectory = new AzureActiveDirectory.ProvisioningAzureActiveDirectory();
            this.Drive = new Drive.Drive();
            this.ProvisioningWebhooks = new ProvisioningWebhookCollection(null);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Any parameters that can be used throughout the template
        /// </summary>
        public Dictionary<string, string> Parameters { get; internal set; } = new Dictionary<string, string>();

        /// <summary>
        /// Gets or sets the Localizations
        /// </summary>
        public LocalizationCollection Localizations { get; internal set; }

        /// <summary>
        /// The Tenant-wide settings for the template
        /// </summary>
        public ProvisioningTenant Tenant { get; set; }

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
        /// The Connector which can be used to retrieve or save related artifacts
        /// </summary>
        public FileConnectorBase Connector { get; set; }

        /// <summary>
        /// A collection of Provisioning Template objects, if any
        /// </summary>
        public ProvisioningTemplateCollection Templates { get; private set; }

        /// <summary>
        /// A collection of Provisioning Sequence objects, if any
        /// </summary>
        public ProvisioningSequenceCollection Sequences { get; private set; }

        /// <summary>
        /// Settings for provisioning Teams objects, if any
        /// </summary>
        public Teams.ProvisioningTeams Teams { get; private set; }

        /// <summary>
        /// Settings for provisioning Azure Active Directory objects, if any
        /// </summary>
        public AzureActiveDirectory.ProvisioningAzureActiveDirectory AzureActiveDirectory { get; private set; }

        /// <summary>
        /// Settings for provisioning Drive objects, if any
        /// </summary>
        public Drive.Drive Drive { get; private set; }

        /// <summary>
        /// A collection of Provisioning Webhooks
        /// </summary>
        public ProvisioningWebhookCollection ProvisioningWebhooks { get; private set; }


        #endregion
    }
}
