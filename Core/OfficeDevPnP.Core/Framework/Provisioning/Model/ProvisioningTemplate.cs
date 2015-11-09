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
        #region Private Members

        private Dictionary<string, string> _parameters = new Dictionary<string, string>();
        private List<Field> _siteFields = new List<Field>();
        private List<ContentType> _contentTypes = new List<ContentType>();
        private List<PropertyBagEntry> _propertyBags = new List<PropertyBagEntry>();
        private List<ListInstance> _lists = new List<ListInstance>();
        private ComposedLook _composedLook = new ComposedLook();
        private Features _features = new Features();
        private SiteSecurity _siteSecurity = new SiteSecurity();
        private CustomActions _customActions = new CustomActions();
        private List<File> _files = new List<File>();
        private List<Provider> _providers = new List<Provider>();
        private List<Page> _pages = new List<Page>();
        private List<TermGroup> _termGroups = new List<TermGroup>();
        private List<Localization> _siteFieldsLocalization = new List<Localization>();
        private FileConnectorBase connector;
        private string _id;

        private RegionalSettings _regionalSettings = null;
        private List<SupportedUILanguage> _supportedUILanguages = new List<SupportedUILanguage>();
        private AuditSettings _auditSettings = null;
        private Workflows _workflows = null;
        private List<AddIn> _addins = new List<AddIn>();
        private Publishing _publishing = null;
        private Dictionary<String, String> _properties = new Dictionary<string, string>();

        #endregion

        #region Constructors

        public ProvisioningTemplate()
        {
            this.connector = new FileSystemConnector(".", "");
        }

        public ProvisioningTemplate(FileConnectorBase connector)
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

        public List<PropertyBagEntry> PropertyBagEntries
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
            set { this._siteSecurity = value; }
        }

        /// <summary>
        /// Gets a collection of fields 
        /// </summary>
        public List<Field> SiteFields
        {
            get { return this._siteFields; }
            private set { this._siteFields = value; }
        }

        /// <summary>
        /// Gets a collection of Localizations for Site fields
        /// </summary>
        public List<Localization> SiteFieldsLocalizations
        {
            get { return this._siteFieldsLocalization; }
            private set { this._siteFieldsLocalization = value; }
        }

        /// <summary>
        /// Gets a collection of Content Types to create
        /// </summary>
        public List<ContentType> ContentTypes
        {
            get { return this._contentTypes; }
            private set { this._contentTypes = value; }
        }

        public List<ListInstance> Lists
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
            set { this._features = value; }
        }

        /// <summary>
        /// Gets or sets CustomActions for the template
        /// </summary>
        public CustomActions CustomActions
        {
            get { return this._customActions; }
            set { this._customActions = value; }
        }

        /// <summary>
        /// Gets a collection of files for the template
        /// </summary>
        public List<File> Files
        {
            get { return this._files; }
            private set { this._files = value; }
        }

        /// <summary>
        /// Gets or Sets the composed look of the template
        /// </summary>
        public ComposedLook ComposedLook
        {
            get { return this._composedLook; }
            set { this._composedLook = value; }
        }

        /// <summary>
        /// Gets a collection of Providers that are used during the extensibility pipeline
        /// </summary>
        public List<Provider> Providers
        {
            get { return this._providers; }
            private set { this._providers = value; }
        }

        /// <summary>
        /// Gets a collection of Wiki Pages for the template
        /// </summary>
        public List<Page> Pages
        {
            get { return this._pages; }
            private set { this._pages = value; }
        }

        /// <summary>
        /// Gets a collection of termgroups to deploy to the site
        /// </summary>
        public List<TermGroup> TermGroups
        {
            get { return this._termGroups; }
            private set { this._termGroups = value; }
        }

        /// <summary>
        /// The Regional Settings of the Provisioning Template
        /// </summary>
        public RegionalSettings RegionalSettings
        {
            get { return this._regionalSettings; }
            set { this._regionalSettings = value; }
        }

        /// <summary>
        /// The Supported UI Languages for the Provisioning Template
        /// </summary>
        public List<SupportedUILanguage> SupportedUILanguages
        {
            get { return this._supportedUILanguages; }
            private set { this._supportedUILanguages = value; }
        }

        /// <summary>
        /// The Audit Settings for the Provisioning Template
        /// </summary>
        public AuditSettings AuditSettings
        {
            get { return this._auditSettings; }
            set { this._auditSettings = value; }
        }

        /// <summary>
        /// Defines the Workflows to provision
        /// </summary>
        public Workflows Workflows
        {
            get { return this._workflows; }
            set { this._workflows = value; }
        }

        /// <summary>
        /// The Search Settings for the Provisioning Template
        /// </summary>
        public String SearchSettings { get; set; }

        /// <summary>
        /// Defines the SharePoint Add-ins to provision
        /// </summary>
        public List<AddIn> AddIns
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
            set { this._publishing = value; }
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

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}|{21}|{22}|{23}|{24}|{25}|{26}|",
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
                this.Providers.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
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
                (this.Publishing != null ? this.Publishing.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningTemplate))
            {
                return (false);
            }
            return (Equals((ProvisioningTemplate)obj));
        }

        public bool Equals(ProvisioningTemplate other)
        {
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
                this.Providers.DeepEquals(other.Providers) &&
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
                this.Publishing == other.Publishing
            );
        }

        #endregion
        /// <summary>
        /// Serializes a template to XML
        /// </summary>
        /// <param name="formatter"></param>
        /// <returns></returns>
        public string ToXML(ITemplateFormatter formatter = null)
        {
            formatter = formatter ?? new XMLPnPSchemaFormatter();
            using (var stream = formatter.ToFormattedTemplate(this))
            {
                return XElement.Load(stream).ToString();
            }
        }
    }
}
