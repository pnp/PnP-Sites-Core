using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration
{
    public partial class ExtractConfiguration
    {

        [JsonIgnore]
        internal ProvisioningTemplate BaseTemplate { get; set; }

        [JsonIgnore]
        public FileConnectorBase FileConnector { get; set; }

        [JsonIgnore]
        public ProvisioningProgressDelegate ProgressDelegate { get; set; }

        [JsonIgnore]
        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        [JsonProperty("persistAssetFiles")]
        public bool PersistAssetFiles { get; set; }

        [JsonProperty("handlers")]
        public List<ConfigurationHandler> Handlers { get; set; } = new List<ConfigurationHandler>();

        [JsonProperty("lists")]
        public Lists.ExtractListsConfiguration Lists { get; set; } = new Lists.ExtractListsConfiguration();

        [JsonProperty("pages")]
        public Pages.ExtractPagesConfiguration Pages { get; set; } = new Pages.ExtractPagesConfiguration();

        [JsonProperty("siteSecurity")]
        public SiteSecurity.ExtractConfiguration SiteSecurity { get; set; } = new SiteSecurity.ExtractConfiguration();

        [JsonProperty("taxonomy")]
        public Taxonomy.ExtractTaxonomyConfiguration Taxonomy { get; set; } = new Taxonomy.ExtractTaxonomyConfiguration();

        [JsonProperty("navigation")]
        public Navigation.ExtractNavigationConfiguration Navigation { get; set; } = new Navigation.ExtractNavigationConfiguration();

        [JsonProperty("siteFooter")]
        public SiteFooter.ExtractSiteFooterConfiguration SiteFooter { get; set; } = new SiteFooter.ExtractSiteFooterConfiguration();

        [JsonProperty("contentTypes")]
        public ContentTypes.ExtractContentTypeConfiguration ContentTypes { get; set; } = new ContentTypes.ExtractContentTypeConfiguration();

        [JsonProperty("searchSettings")]
        public SearchSettings.ExtractSearchConfiguration SearchSettings { get; set; } = new SearchSettings.ExtractSearchConfiguration();

        [JsonProperty("extensibility")]
        public Extensibility.ExtractExtensibilityConfiguration Extensibility { get; set; } = new Extensibility.ExtractExtensibilityConfiguration();

        /// <summary>
        /// Defines Tenant Extraction Settings
        /// </summary>
        [JsonProperty("tenant")]
        public Tenant.ExtractTenantConfiguration Tenant { get; set; } = new Tenant.ExtractTenantConfiguration();

        [JsonProperty("propertyBag")]
        public PropertyBag.ExtractPropertyBagConfiguration PropertyBag { get; set; } = new PropertyBag.ExtractPropertyBagConfiguration();

        [JsonProperty("multiLanguage")]
        public MultiLanguage.ExtractMultiLanguageConfiguration MultiLanguage { get; set; } = new MultiLanguage.ExtractMultiLanguageConfiguration();

        [JsonProperty("publishing")]
        public Publishing.ExtractPublishingConfiguration Publishing { get; set; } = new Publishing.ExtractPublishingConfiguration();

        public static ExtractConfiguration FromCreationInformation(ProvisioningTemplateCreationInformation information)
        {
            var config = new ExtractConfiguration();

            config.BaseTemplate = information.BaseTemplate;
            config.ContentTypes.Groups = information.ContentTypeGroupsToInclude;
            config.Extensibility.Handlers = information.ExtensibilityHandlers;
            config.FileConnector = information.FileConnector;
            if (information.HandlersToProcess == Model.Handlers.All)
            {
                config.Handlers = new List<ConfigurationHandler>();
            }
            else
            {
                foreach (var handler in (Handlers[])Enum.GetValues(typeof(Handlers)))
                {
                    if (information.HandlersToProcess.HasFlag(handler))
                    {
                        if (Enum.TryParse<ConfigurationHandler>(handler.ToString(), out ConfigurationHandler configurationHandler))
                        {
                            config.Handlers.Add(configurationHandler);
                        }
                    }
                }
            }

            config.Pages.IncludeAllClientSidePages = information.IncludeAllClientSidePages;
            config.Taxonomy.IncludeAllTermGroups = information.IncludeAllTermGroups;
            config.Taxonomy.IncludeSiteCollectionTermGroup = information.IncludeSiteCollectionTermGroup;
            config.SiteSecurity.IncludeSiteGroups = information.IncludeSiteGroups;
            config.Taxonomy.IncludeSecurity = information.IncludeTermGroupsSecurity;
            if (information.ListsToExtract != null && information.ListsToExtract.Any())
            {
                foreach (var list in information.ListsToExtract)
                {
                    config.Lists.Lists.Add(new Configuration.Lists.Lists.ExtractListsListsConfiguration()
                    {
                        Title = list
                    });
                }
            }
            if (information.MessagesDelegate != null)
            {
                config.MessagesDelegate = (message, type) =>
                {
                    information.MessagesDelegate(message, type);
                };
            }
            config.PersistAssetFiles = information.PersistBrandingFiles || information.PersistPublishingFiles;
            config.MultiLanguage.PersistResources = information.PersistMultiLanguageResources;
            if (information.ProgressDelegate != null)
            {
                config.ProgressDelegate = (message, step, total) =>
                {
                    information.ProgressDelegate(message, step, total);
                };
            }
            config.PropertyBag.ValuesToPreserve = information.PropertyBagPropertiesToPreserve;
            config.MultiLanguage.ResourceFilePrefix = information.ResourceFilePrefix;
            config.Publishing.Persist = information.PersistPublishingFiles;
            config.Publishing.IncludeNativePublishingFiles = information.IncludeNativePublishingFiles;
            config.SearchSettings.Include = information.IncludeSearchConfiguration;
            return config;
        }

        /// <summary>
        /// Converts the Configuration to a ProvisioningTemplateCreationInformation object for backwards compatibility
        /// </summary>
        /// <param name="web"></param>
        /// <returns></returns>
        public ProvisioningTemplateCreationInformation ToCreationInformation(Web web)
        {

            var ci = new ProvisioningTemplateCreationInformation(web);

            ci.ExtractConfiguration = this;

            ci.PersistBrandingFiles = PersistAssetFiles;
            ci.PersistPublishingFiles = PersistAssetFiles;
            ci.BaseTemplate = web.GetBaseTemplate();
            ci.FileConnector = this.FileConnector;
            ci.IncludeAllClientSidePages = this.Pages.IncludeAllClientSidePages;
            ci.IncludeHiddenLists = this.Lists.IncludeHiddenLists;
            ci.IncludeSiteGroups = this.SiteSecurity.IncludeSiteGroups;
            ci.ContentTypeGroupsToInclude = this.ContentTypes.Groups;
            ci.IncludeContentTypesFromSyndication = !this.ContentTypes.ExcludeFromSyndication;
            ci.IncludeTermGroupsSecurity = this.Taxonomy.IncludeSecurity;
            ci.IncludeSiteCollectionTermGroup = this.Taxonomy.IncludeSiteCollectionTermGroup;
            ci.IncludeSearchConfiguration = this.SearchSettings.Include;
            ci.IncludeAllTermGroups = this.Taxonomy.IncludeAllTermGroups;
            ci.ExtensibilityHandlers = this.Extensibility.Handlers;
            ci.IncludeAllTermGroups = this.Taxonomy.IncludeAllTermGroups;
            ci.IncludeNativePublishingFiles = this.Publishing.IncludeNativePublishingFiles;
            ci.ListsToExtract = this.Lists != null && this.Lists.Lists.Any() ? this.Lists.Lists.Select(l => l.Title).ToList() : null;
            ci.PersistMultiLanguageResources = this.MultiLanguage.PersistResources;
            ci.PersistPublishingFiles = this.Publishing.Persist;
            ci.ResourceFilePrefix = this.MultiLanguage.ResourceFilePrefix;

            if (Handlers.Any())
            {
                ci.HandlersToProcess = Model.Handlers.None;
                foreach (var handler in Handlers)
                {
                    Model.Handlers handlerEnumValue = Model.Handlers.None;
                    switch (handler)
                    {
                        case ConfigurationHandler.Pages:
                            handlerEnumValue = Model.Handlers.PageContents;
                            break;
                        case ConfigurationHandler.Taxonomy:
                            handlerEnumValue = Model.Handlers.TermGroups;
                            break;
                        default:
                            handlerEnumValue = (Model.Handlers)Enum.Parse(typeof(Model.Handlers), handler.ToString());
                            break;
                    }
                    ci.HandlersToProcess |= handlerEnumValue;
                }
            }
            else
            {
                ci.HandlersToProcess = Model.Handlers.All;
            }

            if (this.ProgressDelegate != null)
            {
                ci.ProgressDelegate = (message, step, total) =>
                {
                    ProgressDelegate(message, step, total);
                };
            }
            if (this.MessagesDelegate != null)
            {
                ci.MessagesDelegate = (message, type) =>
                {
                    MessagesDelegate(message, type);
                };
            }

     
            return ci;
        }

        public static ExtractConfiguration FromString(string input)
        {
            //var assembly = Assembly.GetExecutingAssembly();
            //var resourceName = "OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.extract-configuration.schema.json";

            //using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            //using (StreamReader reader = new StreamReader(stream))
            //{
            //    string result = reader.ReadToEnd();

            //    JsonSchema schema = JsonSchema.Parse(result);

            //    var jobject = JObject.Parse(input);

            //    if(!jobject.IsValid(schema))
            //    {
            //        throw new JsonSerializationException("Configuration is not valid according to schema");
            //    }
            //}

            return JsonConvert.DeserializeObject<ExtractConfiguration>(input);
        }
    }
}
