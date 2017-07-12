using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Handles methods for Provisioning Template Creation Information
    /// </summary>
    public class ProvisioningTemplateCreationInformation
    {
        private ProvisioningTemplate baseTemplate;
        private FileConnectorBase fileConnector;
        private bool persistBrandingFiles = false;
        private bool persistMultiLanguageResourceFiles = false;
        private string resourceFilePrefix = "PnP_Resources";
        private bool includeAllTermGroups = false;
        private bool includeSiteCollectionTermGroup = false;
        private bool includeSiteGroups = false;
        private bool includeTermGroupsSecurity = false;
        private bool includeSearchConfiguration = false;
        private List<String> propertyBagPropertiesToPreserve;
        private List<String> contentTypeGroupsToInclude;
        private bool persistPublishingFiles = false;
        private bool includeNativePublishingFiles = false;
        private bool skipVersionCheck = false;
        private List<ExtensibilityHandler> extensibilityHandlers = new List<ExtensibilityHandler>();
        private Handlers handlersToProcess = Handlers.All;
        private bool includeContentTypesFromSyndication = true;

        /// <summary>
        /// Provisioning Progress Delegate
        /// </summary>
        public ProvisioningProgressDelegate ProgressDelegate { get; set; }

        /// <summary>
        /// Provisioning Messages Delegate
        /// </summary>
        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="web">A SharePoint site or subsite</param>
        public ProvisioningTemplateCreationInformation(Web web)
        {
            this.baseTemplate = web.GetBaseTemplate();
            this.propertyBagPropertiesToPreserve = new List<String>();
            this.contentTypeGroupsToInclude = new List<String>();
        }

        /// <summary>
        /// Base template used to compare against when we're "getting" a template
        /// </summary>
        public ProvisioningTemplate BaseTemplate
        {
            get
            {
                return this.baseTemplate;
            }
            set
            {
                this.baseTemplate = value;
            }
        }

        /// <summary>
        /// Connector used to persist files when needed
        /// </summary>
        public FileConnectorBase FileConnector
        {
            get
            {
                return this.fileConnector;
            }
            set
            {
                this.fileConnector = value;
            }
        }

        /// <summary>
        /// Will create resource files named "PnP_Resource_[LCID].resx for every supported language. The files will be persisted to the location specified by the connector
        /// </summary>
        public bool PersistMultiLanguageResources
        {
            get
            {
                return this.persistMultiLanguageResourceFiles;
            }
            set
            {
                this.persistMultiLanguageResourceFiles = value;
            }
        }

        /// <summary>
        /// Prefix for resource file
        /// </summary>
        public string ResourceFilePrefix
        {
            get
            {
                return this.resourceFilePrefix;
            }
            set
            {
                this.resourceFilePrefix = value;
            }
        }

        /// <summary>
        /// Do composed look files (theme files, site logo, alternate css) need to be persisted to storage when 
        /// we're "getting" a template
        /// </summary>
        [Obsolete("Use PersistBrandingFiles instead")]
        public bool PersistComposedLookFiles
        {
            get
            {
                return this.persistBrandingFiles;
            }
            set
            {
                this.persistBrandingFiles = value;
            }
        }

        /// <summary>
        /// if true, persists branding files in the template
        /// </summary>
        public bool PersistBrandingFiles
        {
            get
            {
                return this.persistBrandingFiles;
            }
            set
            {
                this.persistBrandingFiles = value;
            }
        }

        /// <summary>
        /// Defines whether to persist publishing files (MasterPages and PageLayouts)
        /// </summary>
        public bool PersistPublishingFiles
        {
            get
            {
                return this.persistPublishingFiles;
            }
            set
            {
                this.persistPublishingFiles = value;
            }
        }

        /// <summary>
        /// Defines whether to extract native publishing files (MasterPages and PageLayouts)
        /// </summary>
        public bool IncludeNativePublishingFiles
        {
            get
            {
                return this.includeNativePublishingFiles;
            }
            set
            {
                this.includeNativePublishingFiles = value;
            }
        }
        
        /// <summary>
        /// If true includes all term groups in the template
        /// </summary>
        public bool IncludeAllTermGroups
        {
            get
            {
                return this.includeAllTermGroups;
            }
            set { this.includeAllTermGroups = value; }
        }

        /// <summary>
        /// if true, includes site collection term groups in the template
        /// </summary>
        public bool IncludeSiteCollectionTermGroup
        {
            get { return this.includeSiteCollectionTermGroup; }
            set { this.includeSiteCollectionTermGroup = value; }
        }

        /// <summary>
        /// if true, includes term group security in the template
        /// </summary>
        public bool IncludeTermGroupsSecurity
        {
            get { return this.includeTermGroupsSecurity; }
            set { this.includeTermGroupsSecurity = value; }
        }

        internal List<String> PropertyBagPropertiesToPreserve
        {
            get { return this.propertyBagPropertiesToPreserve; }
            set { this.propertyBagPropertiesToPreserve = value; }
        }

        /// <summary>
        /// List of content type groups
        /// </summary>
        public List<String> ContentTypeGroupsToInclude {
            get { return this.contentTypeGroupsToInclude; }
            set { this.contentTypeGroupsToInclude = value; }
        }

        /// <summary>
        /// if true, includes site groups in the template
        /// </summary>
        public bool IncludeSiteGroups
        {
            get
            {
                return this.includeSiteGroups;
            }
            set { this.includeSiteGroups = value; }
        }

        /// <summary>
        /// if true includes search configuration in the template
        /// </summary>
        public bool IncludeSearchConfiguration
        {
            get
            {
                return this.includeSearchConfiguration;
            }
            set
            {
                this.includeSearchConfiguration = value;
            }
        }

        /// <summary>
        /// List of of handlers to process
        /// </summary>
        public Handlers HandlersToProcess
        {
            get
            {
                return handlersToProcess;
            }
            set
            {
                handlersToProcess = value;
            }
        }

        /// <summary>
        /// List of ExtensibilityHandlers
        /// </summary>
        public List<ExtensibilityHandler> ExtensibilityHandlers
        {
            get
            {
                return extensibilityHandlers;
            }

            set
            {
                extensibilityHandlers = value;
            }
        }

        /// <summary>
        /// if true, skips version check
        /// </summary>
        public bool SkipVersionCheck
        {
            get { return skipVersionCheck; }
            set { skipVersionCheck = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to include content types from syndication (= content type hub) or not.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the export should contains content types issued from syndication (= content type hub)
        /// </value>
        public bool IncludeContentTypesFromSyndication
        {
            get { return includeContentTypesFromSyndication; }
            set { includeContentTypesFromSyndication = value; }
        }
    }
}
