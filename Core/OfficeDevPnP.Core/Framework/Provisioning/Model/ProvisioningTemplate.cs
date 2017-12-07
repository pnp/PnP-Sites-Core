using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.IO;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object for the Provisioning Template
    /// </summary>
    public partial class ProvisioningTemplate : IEquatable<ProvisioningTemplate>
    {
        #region Private Fields

        private Dictionary<string, string> _parameters = new Dictionary<string, string>();
        private LocalizationCollection _localizations;
        private FieldCollection _siteFields;
        private ContentTypeCollection _contentTypes;
        private PropertyBagEntryCollection _propertyBags;
        private ListInstanceCollection _lists;
        private ComposedLook _composedLook;
        private Features _features;
        private SiteSecurity _siteSecurity;
        private Navigation _navigation;
        private CustomActions _customActions;
        private FileCollection _files;
        private DirectoryCollection _directories;
        private ExtensibilityHandlerCollection _extensibilityHandlers;
        private PageCollection _pages;
        private TermGroupCollection _termGroups;
        private FileConnectorBase connector;
        private string _id;

        private RegionalSettings _regionalSettings = null;
        private WebSettings _webSettings = null;
        private SupportedUILanguageCollection _supportedUILanguages;
        private AuditSettings _auditSettings = null;
        private Workflows _workflows = null;
        private AddInCollection _addins;
        private Publishing _publishing = null;
        private Dictionary<String, String> _properties = new Dictionary<string, string>();

        private SiteWebhookCollection _siteWebhooks;
        private ClientSidePageCollection _clientSidePages;

        private ProvisioningTenant _tenant;
        private ApplicationLifecycleManagement _applicationLifecycleManagement;

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for ProvisioningTemplate class
        /// </summary>
        public ProvisioningTemplate()
        {
            this.connector = new FileSystemConnector(".", "");

            this._localizations = new LocalizationCollection(this);
            this._siteFields = new FieldCollection(this);
            this._contentTypes = new ContentTypeCollection(this);
            this._propertyBags = new PropertyBagEntryCollection(this);
            this._lists = new ListInstanceCollection(this);

            this._siteSecurity = new SiteSecurity();
            this._siteSecurity.ParentTemplate = this;

            this._composedLook = new ComposedLook();
            this._composedLook.ParentTemplate = this;
            this._features = new Features();
            this._features.ParentTemplate = this;
            this._customActions = new CustomActions();
            this._customActions.ParentTemplate = this;

            this._files = new FileCollection(this);
            this._directories = new DirectoryCollection(this);
            this._providers = new ProviderCollection(this); // Deprecated
            this._extensibilityHandlers = new ExtensibilityHandlerCollection(this);
            this._pages = new PageCollection(this);
            this._termGroups = new TermGroupCollection(this);

            this._supportedUILanguages = new SupportedUILanguageCollection(this);
            this._addins = new AddInCollection(this);

            this._siteWebhooks = new SiteWebhookCollection(this);
            this._clientSidePages = new ClientSidePageCollection(this);

            this._tenant = new ProvisioningTenant();
            this._tenant.ParentTemplate = this;

            this._applicationLifecycleManagement = new ApplicationLifecycleManagement();
            this._applicationLifecycleManagement.ParentTemplate = this;
        }

        /// <summary>
        /// Constructor for ProvisioningTemplate class
        /// </summary>
        /// <param name="connector">FileConnectorBase object</param>
        public ProvisioningTemplate(FileConnectorBase connector) :
            this()
        {
            this.connector = connector;
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
        /// Gets or sets the ID of the Provisioning Template
        /// </summary>
        public string Id { get { return _id; } set { _id = value; } }

        /// <summary>
        /// Gets or sets the Version of the Provisioning Template
        /// </summary>
        public double Version { get; set; }

        /// <summary>
        /// Gets or Sets the Site Policy
        /// </summary>
        public string SitePolicy { get; set; }

        public PropertyBagEntryCollection PropertyBagEntries
        {
            get { return this._propertyBags; }
            private set { this._propertyBags = value; }
        }

        /// <summary>
        /// Security Groups Members for the Template
        /// </summary>
        public SiteSecurity Security
        {
            get { return this._siteSecurity; }
            set
            {
                if (this._siteSecurity != null)
                {
                    this._siteSecurity.ParentTemplate = null;
                }
                this._siteSecurity = value;
                if (this._siteSecurity != null)
                {
                    this._siteSecurity.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// The Navigation configurations of the Provisioning Template
        /// </summary>
        public Navigation Navigation
        {
            get { return this._navigation; }
            set
            {
                if (this._navigation != null)
                {
                    this._navigation.ParentTemplate = null;
                }
                this._navigation = value;
                if (this._navigation != null)
                {
                    this._navigation.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// Gets a collection of fields 
        /// </summary>
        public FieldCollection SiteFields
        {
            get { return this._siteFields; }
            private set { this._siteFields = value; }
        }

        /// <summary>
        /// Gets a collection of Content Types to create
        /// </summary>
        public ContentTypeCollection ContentTypes
        {
            get { return this._contentTypes; }
            private set { this._contentTypes = value; }
        }

        public ListInstanceCollection Lists
        {
            get { return this._lists; }
            private set { this._lists = value; }
        }

        /// <summary>
        /// Gets or sets a list of features to activate or deactivate
        /// </summary>
        public Features Features
        {
            get { return this._features; }
            set
            {
                if (this._features != null)
                {
                    this._features.ParentTemplate = null;
                }
                this._features = value;
                if (this._features != null)
                {
                    this._features.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// Gets or sets CustomActions for the template
        /// </summary>
        public CustomActions CustomActions
        {
            get { return this._customActions; }
            set
            {
                if (this._customActions != null)
                {
                    this._customActions.ParentTemplate = null;
                }
                this._customActions = value;
                if (this._customActions != null)
                {
                    this._customActions.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// Gets a collection of files for the template
        /// </summary>
        public FileCollection Files
        {
            get { return this._files; }
            private set { this._files = value; }
        }

        /// <summary>
        /// Gets a collection of directories from which upload files for the template
        /// </summary>
        public DirectoryCollection Directories
        {
            get { return this._directories; }
            private set { this._directories = value; }
        }

        /// <summary>
        /// Gets or Sets the composed look of the template
        /// </summary>
        public ComposedLook ComposedLook
        {
            get { return this._composedLook; }
            set
            {
                if (this._composedLook != null)
                {
                    this._composedLook.ParentTemplate = null;
                }
                this._composedLook = value;
                if (this._composedLook != null)
                {
                    this._composedLook.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// Gets or sets the Extensibility Handlers
        /// </summary>
        public ExtensibilityHandlerCollection ExtensibilityHandlers
        {
            get { return this._extensibilityHandlers; }
            private set { this._extensibilityHandlers = value; }
        }


        /// <summary>
        /// Gets a collection of Wiki Pages for the template
        /// </summary>
        public PageCollection Pages
        {
            get { return this._pages; }
            private set { this._pages = value; }
        }

        /// <summary>
        /// Gets a collection of termgroups to deploy to the site
        /// </summary>
        public TermGroupCollection TermGroups
        {
            get { return this._termGroups; }
            private set { this._termGroups = value; }
        }

        /// <summary>
        /// The Web Settings of the Provisioning Template
        /// </summary>
        public WebSettings WebSettings
        {
            get { return this._webSettings; }
            set
            {
                if (this._webSettings != null)
                {
                    this._webSettings.ParentTemplate = null;
                }
                this._webSettings = value;
                if (this._webSettings != null)
                {
                    this._webSettings.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// The Regional Settings of the Provisioning Template
        /// </summary>
        public RegionalSettings RegionalSettings
        {
            get { return this._regionalSettings; }
            set
            {
                if (this._regionalSettings != null)
                {
                    this._regionalSettings.ParentTemplate = null;
                }
                this._regionalSettings = value;
                if (this._regionalSettings != null)
                {
                    this._regionalSettings.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// The Supported UI Languages for the Provisioning Template
        /// </summary>
        public SupportedUILanguageCollection SupportedUILanguages
        {
            get { return this._supportedUILanguages; }
            internal set { this._supportedUILanguages = value; }
        }

        /// <summary>
        /// The Audit Settings for the Provisioning Template
        /// </summary>
        public AuditSettings AuditSettings
        {
            get
            {
                return this._auditSettings;
            }
            set
            {
                // If we already have an AuditSettings bounded
                if (this._auditSettings != null)
                {
                    // Clear its parent template
                    this._auditSettings.ParentTemplate = null;
                }
                // Set the new AuditSettings instance
                this._auditSettings = value;
                if (this._auditSettings != null)
                {
                    // Make this template as the parent template of the new AuditSettings instance
                    this._auditSettings.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// Defines the Workflows to provision
        /// </summary>
        public Workflows Workflows
        {
            get { return this._workflows; }
            set
            {
                if (this._workflows != null)
                {
                    this._workflows.ParentTemplate = null;
                }
                this._workflows = value;
                if (this._workflows != null)
                {
                    this._workflows.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// The Site Collection level Search Settings for the Provisioning Template
        /// </summary>
        public String SiteSearchSettings { get; set; }

        /// <summary>
        /// The Web level Search Settings for the Provisioning Template
        /// </summary>
        public String WebSearchSettings { get; set; }

        /// <summary>
        /// Defines the SharePoint Add-ins to provision
        /// </summary>
        public AddInCollection AddIns
        {
            get { return this._addins; }
            private set { this._addins = value; }
        }

        /// <summary>
        /// Defines the Publishing configuration to provision
        /// </summary>
        public Publishing Publishing
        {
            get { return this._publishing; }
            set
            {
                if (this._publishing != null)
                {
                    this._publishing.ParentTemplate = null;
                }
                this._publishing = value;
                if (this._publishing != null)
                {
                    this._publishing.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// Gets a collection of SiteWebhooks to configure for the site
        /// </summary>
        public SiteWebhookCollection SiteWebhooks
        {
            get { return this._siteWebhooks; }
            private set { this._siteWebhooks = value; }
        }

        /// <summary>
        /// Gets a collection of ClientSidePage to configure for the site
        /// </summary>
        public ClientSidePageCollection ClientSidePages
        {
            get { return this._clientSidePages; }
            private set { this._clientSidePages = value; }
        }

        /// <summary>
        /// A set of custom Properties for the Provisioning Template
        /// </summary>
        public Dictionary<String, String> Properties
        {
            get { return this._properties; }
            private set { this._properties = value; }
        }

        /// <summary>
        /// The Tenant-wide settings for the template
        /// </summary>
        public ProvisioningTenant Tenant
        {
            get { return this._tenant; }
            set
            {
                if (this._tenant != null)
                {
                    this._tenant.ParentTemplate = null;
                }
                this._tenant = value;
                if (this._tenant != null)
                {
                    this._tenant.ParentTemplate = this;
                }
            }
        }

        public ApplicationLifecycleManagement ApplicationLifecycleManagement
        {
            get { return this._applicationLifecycleManagement; }
            set
            {
                if (this._applicationLifecycleManagement != null)
                {
                    this._applicationLifecycleManagement.ParentTemplate = null;
                }
                this._applicationLifecycleManagement = value;
                if (this._applicationLifecycleManagement != null)
                {
                    this._applicationLifecycleManagement.ParentTemplate = this;
                }
            }
        }

        /// <summary>
        /// The Image Preview Url of the Provisioning Template
        /// </summary>
        public String ImagePreviewUrl { get; set; }

        /// <summary>
        /// The Display Name of the Provisioning Template
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// The Description of the Provisioning Template
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The Base SiteTemplate of the Provisioning Template
        /// </summary>
        public String BaseSiteTemplate { get; set; }

        /// <summary>
        /// The default CultureInfo of the Provisioning Template, used to format all input values, optional attribute.
        /// </summary>
        public String TemplateCultureInfo { get; set; }

        /// <summary>
        /// Declares the target scope of the current Provisioning Template
        /// </summary>
        public ProvisioningTemplateScope Scope { get; set; }

        /// <summary>
        /// Gets or sets the File Connector
        /// </summary>
        public FileConnectorBase Connector
        {
            get
            {
                return this.connector;
            }
            set
            {
                this.connector = value;
            }
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}|{21}|{22}|{23}|{24}|{25}|{26}|{27}|{28}|{29}|{30}|{31}|{32}|{33}|",
                (this.ComposedLook != null ? this.ComposedLook.GetHashCode() : 0),
                this.ContentTypes.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.CustomActions.SiteCustomActions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.CustomActions.WebCustomActions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Features.SiteFeatures.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Features.WebFeatures.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Files.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                (this.Id != null ? this.Id.GetHashCode() : 0),
                this.Lists.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.PropertyBagEntries.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
#pragma warning disable 618
                this.Providers.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
#pragma warning restore 618
                this.Security.AdditionalAdministrators.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Security.AdditionalMembers.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Security.AdditionalOwners.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Security.AdditionalVisitors.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Security.SiteGroups.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Security.SiteSecurityPermissions.RoleAssignments.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Security.SiteSecurityPermissions.RoleDefinitions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.SiteFields.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                (this.SitePolicy != null ? this.SitePolicy.GetHashCode() : 0),
                this.Version.GetHashCode(),
                this.Pages.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.TermGroups.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Workflows.WorkflowDefinitions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Workflows.WorkflowSubscriptions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.AddIns.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                (this.Publishing != null ? this.Publishing.GetHashCode() : 0),
                this.Localizations.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.WebSettings.GetHashCode(),
                this.SiteWebhooks.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.ClientSidePages.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.TemplateCultureInfo?.GetHashCode() ?? 0,
                this.Scope.GetHashCode(),
                this.Tenant.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ProvisioningTemplate
        /// </summary>
        /// <param name="obj">Object that represents ProvisioningTemplate</param>
        /// <returns>true if the current object is equal to the ProvisioningTemplate</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningTemplate))
            {
                return (false);
            }
            return (Equals((ProvisioningTemplate)obj));
        }

        /// <summary>
        /// Compares ProvisioningTemplate object based on ComposedLook, ContentTypes, CustomActions, SiteFeature, WebFeatures, Files, Id, Lists,
        /// PropertyBagEntries, Providers, Security, SiteFields, SitePolicy, Version, Pages, TermGroups, Workflows, AddIns, Publishing, Loaclizations,
        /// WebSettings, SiteWebhooks, ClientSidePages, and Tenant properties.
        /// </summary>
        /// <param name="other">ProvisioningTemplate object</param>
        /// <returns>true if the ProvisioningTemplate object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ProvisioningTemplate other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.ComposedLook.Equals(other.ComposedLook) &&
                this.ContentTypes.DeepEquals(other.ContentTypes) &&
                this.CustomActions.SiteCustomActions.DeepEquals(other.CustomActions.SiteCustomActions) &&
                this.CustomActions.WebCustomActions.DeepEquals(other.CustomActions.WebCustomActions) &&
                this.Features.SiteFeatures.DeepEquals(other.Features.SiteFeatures) &&
                this.Features.WebFeatures.DeepEquals(other.Features.WebFeatures) &&
                this.Files.DeepEquals(other.Files) &&
                this.Id == other.Id &&
                this.Lists.DeepEquals(other.Lists) &&
                this.PropertyBagEntries.DeepEquals(other.PropertyBagEntries) &&
#pragma warning disable 618
                this.Providers.DeepEquals(other.Providers) &&
#pragma warning restore 618
                this.Security.AdditionalAdministrators.DeepEquals(other.Security.AdditionalAdministrators) &&
                this.Security.AdditionalMembers.DeepEquals(other.Security.AdditionalMembers) &&
                this.Security.AdditionalOwners.DeepEquals(other.Security.AdditionalOwners) &&
                this.Security.AdditionalVisitors.DeepEquals(other.Security.AdditionalVisitors) &&
                this.Security.SiteGroups.DeepEquals(other.Security.SiteGroups) &&
                this.Security.SiteSecurityPermissions.RoleAssignments.DeepEquals(other.Security.SiteSecurityPermissions.RoleAssignments) &&
                this.Security.SiteSecurityPermissions.RoleDefinitions.DeepEquals(other.Security.SiteSecurityPermissions.RoleDefinitions) &&
                this.SiteFields.DeepEquals(other.SiteFields) &&
                this.SitePolicy == other.SitePolicy &&
                this.Version == other.Version &&
                this.Pages.DeepEquals(other.Pages) &&
                this.TermGroups.DeepEquals(other.TermGroups) &&
                ((this.Workflows != null && other.Workflows != null) ? this.Workflows.WorkflowDefinitions.DeepEquals(other.Workflows.WorkflowDefinitions) : true) &&
                ((this.Workflows != null && other.Workflows != null) ? this.Workflows.WorkflowSubscriptions.DeepEquals(other.Workflows.WorkflowSubscriptions) : true) &&
                this.AddIns.DeepEquals(other.AddIns) &&
                this.Publishing == other.Publishing &&
                this.Localizations.DeepEquals(other.Localizations) &&
                this.WebSettings.Equals(other.WebSettings) &&
                this.SiteWebhooks.DeepEquals(other.SiteWebhooks) &&
                this.ClientSidePages.DeepEquals(other.ClientSidePages) &&
                this.TemplateCultureInfo == other.TemplateCultureInfo &&
                this.Scope == other.Scope &&
                this.Tenant == other.Tenant
            );
        }

        #endregion

        /// <summary>
        /// Serializes a template to XML
        /// </summary>
        /// <param name="formatter">ITemplateFormatter object</param>
        /// <returns>Returns XML string for the given stream</returns>
        public string ToXML(ITemplateFormatter formatter = null)
        {
            formatter = formatter ?? new XMLPnPSchemaFormatter();
            using (var stream = formatter.ToFormattedTemplate(this))
            {
                return XElement.Load(stream).ToString();
            }
        }
    }

    /// <summary>
    /// Declares the target scope of the current Provisioning Template
    /// </summary>
    public enum ProvisioningTemplateScope
    {
        /// <summary>
        /// Value for when scope was not set in the template
        /// </summary>
        Undefined,
        /// <summary>
        /// The scope is a Root web of a Site Collection
        /// </summary>
        RootSite,
        /// <summary>
        /// The scope is a child Web of a Site Collection
        /// </summary>
        Web,
    }
}
