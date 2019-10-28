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
        public FileConnectorBase FileConnector { get; set; }

        [JsonIgnore]
        public Action<string, int, int> ProgressAction { get; set; }

        [JsonIgnore]
        public Action<string, ProvisioningMessageType> MessageAction { get; set; }

        [JsonProperty("persistAssetFiles")]
        public bool PersistAssetFiles { get; set; }

        [JsonProperty("handlers")]
        public List<ConfigurationHandler> Handlers { get; set; } = new List<ConfigurationHandler>();

        [JsonProperty("lists")]
        public Lists.ExtractListsConfiguration Lists { get; set; } = new Lists.ExtractListsConfiguration();

        [JsonProperty("pages")]
        public Pages.ExtractConfiguration Pages { get; set; } = new Pages.ExtractConfiguration();

        [JsonProperty("siteSecurity")]
        public SiteSecurity.ExtractConfiguration SiteSecurity { get; set; } = new SiteSecurity.ExtractConfiguration();

        [JsonProperty("taxonomy")]
        public Taxonomy.ExtractConfiguration Taxonomy { get; set; } = new Taxonomy.ExtractConfiguration();

        [JsonProperty("navigation")]
        public Navigation.ExtractNavigationConfiguration Navigation { get; set; } = new Navigation.ExtractNavigationConfiguration();

        [JsonProperty("siteFooter")]
        public SiteFooter.ExtractConfiguration SiteFooter { get; set; } = new SiteFooter.ExtractConfiguration();

        [JsonProperty("contentTypes")]
        public ContentTypes.ExtractContentTypeConfiguration ContentTypes { get; set; } = new ContentTypes.ExtractContentTypeConfiguration();

        [JsonProperty("searchSettings")]
        public SearchSettings.ExtractConfiguration SearchSettings { get; set; } = new SearchSettings.ExtractConfiguration();

        /// <summary>
        /// Defines Tenant Extraction Settings
        /// </summary>
        [JsonProperty("tenant")]
        public Tenant.ExtractTenantConfiguration Tenant { get; set; } = new Tenant.ExtractTenantConfiguration();

        public ProvisioningTemplateCreationInformation ToCreationInformation(Web web)
        {

            var ci = new ProvisioningTemplateCreationInformation(web);

            ci.ExtractConfiguration = this;

            ci.PersistBrandingFiles = PersistAssetFiles;

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
            ci.FileConnector = this.FileConnector;
            ci.IncludeAllClientSidePages = this.Pages.IncludeAllClientSidePages;
            ci.IncludeHiddenLists = this.Lists.IncludeHiddenLists;
            ci.IncludeSiteGroups = this.SiteSecurity.IncludeSiteGroups;
            ci.ContentTypeGroupsToInclude = this.ContentTypes.Groups;
            ci.IncludeContentTypesFromSyndication = !this.ContentTypes.ExcludeFromSyndication;
            ci.IncludeTermGroupsSecurity = this.Taxonomy.IncludeSecurity;
            ci.IncludeSiteCollectionTermGroup = this.Taxonomy.IncludeSiteCollectionTermGroup;
            ci.IncludeSearchConfiguration = this.SearchSettings.Include;

            if (this.ProgressAction != null)
            {
                ci.ProgressDelegate = (message, step, total) =>
                {
                    ProgressAction(message, step, total);
                };
            }
            if (this.MessageAction != null)
            {
                ci.MessagesDelegate = (message, type) =>
                {
                    MessageAction(message, type);
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
