using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
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
        [JsonProperty("persistAssetFiles")]
        public bool PersistAssetFiles { get; set; }

        [JsonProperty("handlers")]
        public List<ConfigurationHandler> Handlers { get; set; }

        [JsonProperty("lists")]
        public Lists.ExtractConfiguration Lists { get; set; }

        [JsonProperty("pages")]
        public Pages.ExtractConfiguration Pages { get; set; }

        [JsonProperty("siteSecurity")]
        public SiteSecurity.ExtractConfiguration SiteSecurity { get; set; }

        [JsonProperty("taxonomy")]
        public Taxonomy.ExtractConfiguration Taxonomy { get; set; }

        [JsonProperty("navigation")]
        public Navigation.ExtractConfiguration Navigation { get; set; }

        [JsonProperty("siteFooter")]
        public SiteFooter.ExtractConfiguration SiteFooter { get; set; }

        [JsonProperty("contentTypes")]
        public ContentTypes.ExtractConfiguration ContentTypes { get; set; }

        [JsonProperty("searchSettings")]
        public SearchSettings.ExtractConfiguration SearchSettings
        {
            get; set;
        }
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

            if (this.Pages != null)
            {
                ci.IncludeAllClientSidePages = this.Pages.IncludeAllClientSidePages;
            }
            if (this.Lists != null)
            {
                ci.IncludeHiddenLists = this.Lists.IncludeHiddenLists;
            }
            if (this.SiteSecurity != null)
            {
                ci.IncludeSiteGroups = this.SiteSecurity.IncludeSiteGroups;
            }
            if (this.ContentTypes != null)
            {
                ci.ContentTypeGroupsToInclude = this.ContentTypes.Groups;
                ci.IncludeContentTypesFromSyndication = !this.ContentTypes.ExcludeFromSyndication;
            }
            if (this.Taxonomy != null)
            {
                ci.IncludeTermGroupsSecurity = this.Taxonomy.IncludeSecurity;
                ci.IncludeSiteCollectionTermGroup = this.Taxonomy.IncludeSiteCollectionTermGroup;
            }
            if (this.SearchSettings != null)
            {
                ci.IncludeSearchConfiguration = this.SearchSettings.Include;
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
