using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
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

        public ProvisioningProgressDelegate ProgressDelegate { get; set; }
        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

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
        
        public bool IncludeAllTermGroups
        {
            get
            {
                return this.includeAllTermGroups;
            }
            set { this.includeAllTermGroups = value; }
        }

        public bool IncludeSiteCollectionTermGroup
        {
            get { return this.includeSiteCollectionTermGroup; }
            set { this.includeSiteCollectionTermGroup = value; }
        }

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

        public List<String> ContentTypeGroupsToInclude {
            get { return this.contentTypeGroupsToInclude; }
            set { this.contentTypeGroupsToInclude = value; }
        }

        public bool IncludeSiteGroups
        {
            get
            {
                return this.includeSiteGroups;
            }
            set { this.includeSiteGroups = value; }
        }

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

        public bool SkipVersionCheck
        {
            get { return skipVersionCheck; }
            set { skipVersionCheck = value; }
        }
    }
}
