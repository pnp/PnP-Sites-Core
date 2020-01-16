using AngleSharp.Dom;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net;
using System.Text;
#if !NETSTANDARD2_0
using System.Web.UI;
#endif

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// This class is used to instantiate controls of type 3 (= client side web parts). Using this class you can instantiate a control and 
    /// add it on a <see cref="ClientSidePage"/>.
    /// </summary>
    public class ClientSideWebPart : CanvasControl
    {
        #region variables
        // Constants
        public const string WebPartAttribute = "data-sp-webpart";
        public const string WebPartDataVersionAttribute = "data-sp-webpartdataversion";
        public const string WebPartDataAttribute = "data-sp-webpartdata";
        public const string WebPartComponentIdAttribute = "data-sp-componentid";
        public const string WebPartHtmlPropertiesAttribute = "data-sp-htmlproperties";

        private ClientSideComponent component;
        private string jsonWebPartData;
        private string htmlPropertiesData;
        private string htmlProperties;
        private string webPartId;
        private string webPartData;
        private string title;
        private bool supportsFullBleed;
        private string description;
        private string propertiesJson;
        private ClientSideWebPartControlData spControlData;
        private JObject properties;
        private JObject serverProcessedContent;
        private JObject dynamicDataPaths;
        private JObject dynamicDataValues;
        private string webPartPreviewImage;
        private bool usingSpControlDataOnly;
        #endregion

        #region construction
        /// <summary>
        /// Instantiates client side web part from scratch.
        /// </summary>
        public ClientSideWebPart() : base()
        {
            this.controlType = 3;
            this.webPartData = "";
            this.htmlPropertiesData = "";
            this.htmlProperties = "";
            this.title = "";
            this.description = "";
            this.supportsFullBleed = false;
            this.SetPropertiesJson("{}");
            this.webPartPreviewImage = "";
            this.usingSpControlDataOnly = false;
            this.dynamicDataPaths = JObject.Parse("{}");
            this.dynamicDataValues = JObject.Parse("{}");
            this.serverProcessedContent = JObject.Parse("{}");
        }

        /// <summary>
        /// Instantiates a client side web part based on the information that was obtain from calling the AvailableClientSideComponents methods on the <see cref="ClientSidePage"/> object.
        /// </summary>
        /// <param name="component">Component to create a ClientSideWebPart instance for</param>
        public ClientSideWebPart(ClientSideComponent component) : this()
        {
            if (component == null)
            {
                throw new ArgumentNullException("Passed in component cannot be null");
            }
            this.Import(component);
        }
        #endregion

        #region Properties
        /// <summary>
        /// Value of the "data-sp-webpartdata" attribute
        /// </summary>
        public string JsonWebPartData
        {
            get
            {
                return jsonWebPartData;
            }
        }

        /// <summary>
        /// Value of the "data-sp-htmlproperties" element
        /// </summary>
        public string HtmlPropertiesData
        {
            get
            {
                return htmlPropertiesData;
            }
        }

        /// <summary>
        /// Value of the "data-sp-htmlproperties" attribute
        /// </summary>
        public string HtmlProperties
        {
            get
            {
                return htmlProperties;
            }

        }

        /// <summary>
        /// ID of the client side web part
        /// </summary>
        public string WebPartId
        {
            get
            {
                return webPartId;
            }
        }

        /// <summary>
        /// Supports full bleed display experience
        /// </summary>
        public bool SupportsFullBleed
        {
            get
            {
                return supportsFullBleed;
            }
        }

        /// <summary>
        /// Value of the "data-sp-webpart" attribute
        /// </summary>
        public string WebPartData
        {
            get
            {
                return webPartData;
            }
        }

        /// <summary>
        /// Title of the web part
        /// </summary>
        public string Title
        {
            get
            {
                return this.title;
            }
            set
            {
                this.title = value;
            }

        }

        /// <summary>
        /// Description of the web part
        /// </summary>
        public string Description
        {
            get
            {
                return this.description;
            }
            set
            {
                this.description = value;
            }
        }

        /// <summary>
        /// Preview image that can serve as page preview image when the page holding this web part is promoted to a news page
        /// </summary>
        public string WebPartPreviewImage
        {
            get
            {
                return this.webPartPreviewImage;
            }
        }

        /// <summary>
        /// Json serialized web part information. For 1st party web parts this ideally is the *full* JSON string 
        /// fetch via workbench or via copying it from an existing page. It's important that the serverProcessedContent
        /// element is included here!
        /// </summary>
        public string PropertiesJson
        {
            get
            {
                return this.Properties.ToString(Formatting.None);
            }
            set
            {
                this.SetPropertiesJson(value);
            }
        }

        /// <summary>
        /// Web properties as configurable <see cref="JObject"/>
        /// </summary>
        public JObject Properties
        {
            get
            {
                return this.properties;
            }
        }

        /// <summary>
        /// ServerProcessedContent json node
        /// </summary>
        public JObject ServerProcessedContent
        {
            get
            {
                return this.serverProcessedContent;
            }
        }

        public JObject DynamicDataPaths
        {
            get
            {
                return this.dynamicDataPaths;
            }
        }

        public JObject DynamicDataValues
        {
            get
            {
                return this.dynamicDataValues;
            }
        }

        /// <summary>
        /// Return <see cref="Type"/> of the client side web part
        /// </summary>
        public override Type Type
        {
            get
            {
                return typeof(ClientSideWebPart);
            }
        }

        /// <summary>
        /// Value of the "data-sp-controldata" attribute
        /// </summary>
        public ClientSideWebPartControlData SpControlData
        {
            get
            {
                return this.spControlData;
            }
        }

        /// <summary>
        /// Indicates that this control is persisted/read using the data-sp-controldata attribute only
        /// </summary>
        public bool UsingSpControlDataOnly
        {
            get
            {
                return this.usingSpControlDataOnly;
            }
            set
            {
                this.usingSpControlDataOnly = value;
            }
        }
        #endregion

        #region public methods
        /// <summary>
        /// Imports a <see cref="ClientSideComponent"/> to use it as base for configuring the client side web part instance
        /// </summary>
        /// <param name="component"><see cref="ClientSideComponent"/> to import</param>
        /// <param name="clientSideWebPartPropertiesUpdater">Function callback that allows you to manipulate the client side web part properties after import</param>
        public void Import(ClientSideComponent component, Func<String, String> clientSideWebPartPropertiesUpdater = null)
        {
            this.component = component;
            // Sometimes the id guid is encoded with curly brackets, so let's drop those
            this.webPartId = new Guid(component.Id).ToString("D");

            // Parse the manifest json blob as we need some data from it
            JObject wpJObject = JObject.Parse(component.Manifest);
            this.title = wpJObject["preconfiguredEntries"][0]["title"]["default"].Value<string>();
            this.description = wpJObject["preconfiguredEntries"][0]["title"]["default"].Value<string>();
            if (wpJObject["supportsFullBleed"] != null)
            {
                this.supportsFullBleed = wpJObject["supportsFullBleed"].Value<bool>();
            }
            else
            {
                this.supportsFullBleed = false;
            }

            this.SetPropertiesJson(wpJObject["preconfiguredEntries"][0]["properties"].ToString(Formatting.None));

            if (clientSideWebPartPropertiesUpdater != null)
            {
                this.propertiesJson = clientSideWebPartPropertiesUpdater(this.propertiesJson);
            }
        }

        /// <summary>
        /// Returns a HTML representation of the client side web part
        /// </summary>
        /// <param name="controlIndex">The sequence of the control inside the section</param>
        /// <returns>HTML representation of the client side web part</returns>
        public override string ToHtml(float controlIndex)
        {
            // Can this control be hosted in this section type?
            if (this.Section.Type == CanvasSectionTemplate.OneColumnFullWidth)
            {
                if (!this.SupportsFullBleed)
                {
                    throw new Exception("You cannot host this web part inside a one column full width section, only webparts that support full bleed are allowed");
                }
            }

            ClientSideWebPartControlData controlData = null;
            if (this.usingSpControlDataOnly)
            {
                controlData = new ClientSideWebPartControlDataOnly();
            }
            else
            {
                controlData = new ClientSideWebPartControlData();
            }

            // Obtain the json data
            controlData.ControlType = this.ControlType;
            controlData.Id = this.InstanceId.ToString("D");
            controlData.WebPartId = this.WebPartId;
            controlData.Position = new ClientSideCanvasControlPosition()
            {
                ZoneIndex = this.Section.Order,
                SectionIndex = this.Column.Order,
                SectionFactor = this.Column.ColumnFactor,
#if !SP2019
                LayoutIndex = this.Column.LayoutIndex,
#endif
                ControlIndex = controlIndex,
            };

#if !SP2019
            if (this.section.Type == CanvasSectionTemplate.OneColumnVerticalSection)
            {
                if (this.section.Columns.First().Equals(this.Column))
                {
                    controlData.Position.SectionFactor = 12;
                }
            }
#endif

            controlData.Emphasis = new ClientSideSectionEmphasis()
            {
                ZoneEmphasis = this.Column.VerticalSectionEmphasis.HasValue ? this.Column.VerticalSectionEmphasis.Value : this.Section.ZoneEmphasis,
            };

            // Set the control's data version to the latest version...default was 1.0, but some controls use a higher version
            var webPartType = ClientSidePage.NameToClientSideWebPartEnum(controlData.WebPartId);

            // if we read the control from the page then the value might already be set to something different than 1.0...if so, leave as is
            if (this.DataVersion == "1.0")
            {
                if (webPartType == DefaultClientSideWebParts.Image)
                {
                    this.dataVersion = "1.8";
                }
                else if (webPartType == DefaultClientSideWebParts.ImageGallery)
                {
                    this.dataVersion = "1.6";
                }
                else if (webPartType == DefaultClientSideWebParts.People)
                {
                    this.dataVersion = "1.2";
                }
                else if (webPartType == DefaultClientSideWebParts.DocumentEmbed)
                {
                    this.dataVersion = "1.1";
                }
                else if (webPartType == DefaultClientSideWebParts.ContentRollup)
                {
                    this.dataVersion = "2.1";
                }
                else if (webPartType == DefaultClientSideWebParts.QuickLinks)
                {
                    this.dataVersion = "2.0";
                }
            }

            // Set the web part preview image url
            if (this.ServerProcessedContent != null && this.ServerProcessedContent["imageSources"] != null)
            {
                foreach (JProperty property in this.ServerProcessedContent["imageSources"])
                {
                    if (!string.IsNullOrEmpty(property.Value.ToString()))
                    {
                        this.webPartPreviewImage = property.Value.ToString().ToLower();
                        break;
                    }
                }
            }

            ClientSideWebPartData webpartData = new ClientSideWebPartData() { Id = controlData.WebPartId, InstanceId = controlData.Id, Title = this.Title, Description = this.Description, DataVersion = this.DataVersion, Properties = "jsonPropsToReplacePnPRules", DynamicDataPaths = "jsonDynamicDataPathsToReplacePnPRules", DynamicDataValues = "jsonDynamicDataValuesToReplacePnPRules", ServerProcessedContent = "jsonServerProcessedContentToReplacePnPRules" };

            if (this.usingSpControlDataOnly)
            {
                (controlData as ClientSideWebPartControlDataOnly).WebPartData = "jsonWebPartDataToReplacePnPRules";
                this.jsonControlData = JsonConvert.SerializeObject(controlData);
                this.jsonWebPartData = JsonConvert.SerializeObject(webpartData);
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonPropsToReplacePnPRules\"", this.Properties.ToString(Formatting.None));
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonServerProcessedContentToReplacePnPRules\"", this.ServerProcessedContent.ToString(Formatting.None));
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonDynamicDataPathsToReplacePnPRules\"", this.DynamicDataPaths.ToString(Formatting.None));
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonDynamicDataValuesToReplacePnPRules\"", this.DynamicDataValues.ToString(Formatting.None));
                this.jsonControlData = this.jsonControlData.Replace("\"jsonWebPartDataToReplacePnPRules\"", this.jsonWebPartData);
            }
            else
            {
                this.jsonControlData = JsonConvert.SerializeObject(controlData);
                this.jsonWebPartData = JsonConvert.SerializeObject(webpartData);
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonPropsToReplacePnPRules\"", this.Properties.ToString(Formatting.None));
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonServerProcessedContentToReplacePnPRules\"", this.ServerProcessedContent.ToString(Formatting.None));
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonDynamicDataPathsToReplacePnPRules\"", this.DynamicDataPaths.ToString(Formatting.None));
                this.jsonWebPartData = this.jsonWebPartData.Replace("\"jsonDynamicDataValuesToReplacePnPRules\"", this.DynamicDataValues.ToString(Formatting.None));
            }

            StringBuilder html = new StringBuilder(100);
#if NETSTANDARD2_0
            if (this.usingSpControlDataOnly)
            {
                html.Append($@"<div {CanvasControlAttribute}=""{this.CanvasControlData}"" {CanvasDataVersionAttribute}=""{this.DataVersion}"" {ControlDataAttribute}=""{this.JsonControlData.Replace("\"", "&quot;")}""></div>");
            }
            else
            {
                html.Append($@"<div {CanvasControlAttribute}=""{this.CanvasControlData}"" {CanvasDataVersionAttribute}=""{this.DataVersion}"" {ControlDataAttribute}=""{this.JsonControlData.Replace("\"", "&quot;")}"">");
                html.Append($@"<div {WebPartAttribute}=""{this.WebPartData}"" {WebPartDataVersionAttribute}=""{this.DataVersion}"" {WebPartDataAttribute}=""{this.JsonWebPartData.Replace("\"", "&quot;").Replace("<","&lt;").Replace(">","&gt;")}"">");
                html.Append($@"<div {WebPartComponentIdAttribute}="""">");
                html.Append(this.WebPartId);
                html.Append("</div>");
                html.Append($@"<div {WebPartHtmlPropertiesAttribute}=""{this.HtmlProperties}"">");
                RenderHtmlProperties(ref html);
                html.Append("</div>");
                html.Append("</div>");
                html.Append("</div>");
            }
#else
            var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), "");
            try
            {
                if (this.usingSpControlDataOnly)
                {
                    htmlWriter.NewLine = string.Empty;
                    htmlWriter.AddAttribute(CanvasControlAttribute, this.CanvasControlData);
                    htmlWriter.AddAttribute(CanvasDataVersionAttribute, this.CanvasDataVersion);
                    htmlWriter.AddAttribute(ControlDataAttribute, this.JsonControlData);
                    htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);
                    htmlWriter.RenderEndTag();
                }
                else
                {
                    htmlWriter.NewLine = string.Empty;
                    htmlWriter.AddAttribute(CanvasControlAttribute, this.CanvasControlData);
                    htmlWriter.AddAttribute(CanvasDataVersionAttribute, this.CanvasDataVersion);
                    htmlWriter.AddAttribute(ControlDataAttribute, this.JsonControlData);
                    htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);

                    htmlWriter.AddAttribute(WebPartAttribute, this.WebPartData);
                    htmlWriter.AddAttribute(WebPartDataVersionAttribute, this.DataVersion);
                    htmlWriter.AddAttribute(WebPartDataAttribute, this.JsonWebPartData);
                    htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);

                    htmlWriter.AddAttribute(WebPartComponentIdAttribute, "");
                    htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);
                    htmlWriter.Write(this.WebPartId);
                    htmlWriter.RenderEndTag();

                    htmlWriter.AddAttribute(WebPartHtmlPropertiesAttribute, this.HtmlProperties);
                    htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);
                    // Allow for override of the HTML value rendering if this would be needed by controls
                    RenderHtmlProperties(ref htmlWriter);
                    htmlWriter.RenderEndTag();

                    htmlWriter.RenderEndTag();
                    htmlWriter.RenderEndTag();
                }
            }
            finally
            {
                if (htmlWriter != null)
                {
                    htmlWriter.Dispose();
                }
            }
#endif
            return html.ToString();
        }

        /// <summary>
        /// Overrideable method that allows inheriting webparts to control the HTML rendering
        /// </summary>
        /// <param name="htmlWriter">Reference to the html renderer used</param>
#if NETSTANDARD2_0
        protected virtual void RenderHtmlProperties(ref StringBuilder htmlWriter)
        {
            if (this.ServerProcessedContent != null)
            {
                if (this.ServerProcessedContent["searchablePlainTexts"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["searchablePlainTexts"])
                    {
                        htmlWriter.Append($@"<div data-sp-prop-name=""{property.Name}"" data-sp-searchableplaintext=""true"">");
                        htmlWriter.Append(property.Value.ToString());
                        htmlWriter.Append("</div>");
                    }
                }

                if (this.ServerProcessedContent["imageSources"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["imageSources"])
                    {
                        htmlWriter.Append($@"<img data-sp-prop-name=""{property.Name}""");

                        if (!string.IsNullOrEmpty(property.Value.ToString()))
                        {
                            htmlWriter.Append($@" src=""{property.Value.Value<string>()}""");
                        }
                        htmlWriter.Append("></img>");
                    }
                }

                if (this.ServerProcessedContent["links"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["links"])
                    {
                        htmlWriter.Append($@"<a data-sp-prop-name=""{property.Name}"" href=""{property.Value.Value<string>()}""></a>");
                    }
                }

                if (this.ServerProcessedContent["htmlStrings"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["htmlStrings"])
                    {
                        htmlWriter.Append($@"<div data-sp-prop-name=""{property.Name}"">{property.Value.ToString()}</div>");
                    }
                }
            }
            else
            {
                htmlWriter.Append(this.htmlPropertiesData);
            }
        }
#else
        protected virtual void RenderHtmlProperties(ref HtmlTextWriter htmlWriter)
        {
            if (this.ServerProcessedContent != null)
            {
                if (this.ServerProcessedContent["searchablePlainTexts"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["searchablePlainTexts"])
                    {
                        htmlWriter.AddAttribute("data-sp-prop-name", property.Name);
                        htmlWriter.AddAttribute("data-sp-searchableplaintext", "true");
                        htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);
                        htmlWriter.Write(property.Value.ToString());
                        htmlWriter.RenderEndTag();
                    }
                }

                if (this.ServerProcessedContent["imageSources"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["imageSources"])
                    {
                        htmlWriter.AddAttribute("data-sp-prop-name", property.Name);
                        if (!string.IsNullOrEmpty(property.Value.ToString()))
                        {
                            htmlWriter.AddAttribute("src", property.Value.ToString());
                        }
                        htmlWriter.RenderBeginTag(HtmlTextWriterTag.Img);
                        htmlWriter.RenderEndTag();
                    }
                }

                if (this.ServerProcessedContent["links"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["links"])
                    {
                        htmlWriter.AddAttribute("data-sp-prop-name", property.Name);
                        htmlWriter.AddAttribute("href", property.Value.ToString());
                        htmlWriter.RenderBeginTag(HtmlTextWriterTag.A);
                        htmlWriter.RenderEndTag();
                    }
                }

                if (this.ServerProcessedContent["htmlStrings"] != null)
                {
                    foreach (JProperty property in this.ServerProcessedContent["htmlStrings"])
                    {
                        htmlWriter.AddAttribute("data-sp-prop-name", property.Name);
                        htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);
                        htmlWriter.Write(property.Value.ToString());
                        htmlWriter.RenderEndTag();
                    }
                }

            }
            else
            {
                htmlWriter.Write(this.HtmlPropertiesData);
            }
        }
#endif
#endregion

            #region Internal and private methods
        internal override void FromHtml(IElement element)
        {
            base.FromHtml(element);

            // Set/update dataVersion if it was provided as html attribute
            var webPartDataVersion = element.GetAttribute(WebPartDataVersionAttribute);
            if (!string.IsNullOrEmpty(webPartDataVersion))
            {
                this.dataVersion = element.GetAttribute(WebPartDataVersionAttribute);
            }

            // load data from the data-sp-controldata attribute
            var jsonSerializerSettings = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            this.spControlData = JsonConvert.DeserializeObject<ClientSideWebPartControlData>(element.GetAttribute(CanvasControl.ControlDataAttribute), jsonSerializerSettings);
            this.controlType = this.spControlData.ControlType;

            var wpDiv = element.GetElementsByTagName("div").Where(a => a.HasAttribute(ClientSideWebPart.WebPartDataAttribute)).FirstOrDefault();

            // Some components on the page (e.g. web parts configured with isDomainIsolated = true) are rendered differently in the HTML
            if (wpDiv == null)
            {
                JObject wpJObject = JObject.Parse(element.GetAttribute(CanvasControl.ControlDataAttribute));
                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["title"] != null)
                {
                    this.title = wpJObject["webPartData"]["title"].Value<string>();
                }
                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["description"] != null)
                {
                    this.title = wpJObject["webPartData"]["description"].Value<string>();
                }
                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["properties"] != null)
                {
                    this.PropertiesJson = wpJObject["webPartData"]["properties"].ToString(Formatting.None);
                }
                // Check for fullbleed supporting web parts
                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["properties"] != null && wpJObject["webPartData"]["properties"]["isFullWidth"] != null)
                {
                    this.supportsFullBleed = wpJObject["webPartData"]["properties"]["isFullWidth"].Value<Boolean>();
                }

                // Store the server processed content as that's needed for full fidelity
                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["serverProcessedContent"] != null && wpJObject["webPartData"]["serverProcessedContent"].ToString() != "")
                {
                    this.serverProcessedContent = (JObject)wpJObject["webPartData"]["serverProcessedContent"];
                }

                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["dynamicDataPaths"] != null)
                {
                    this.dynamicDataPaths = (JObject)wpJObject["webPartData"]["dynamicDataPaths"];
                }

                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["dynamicDataValues"] != null)
                {
                    this.dynamicDataValues = (JObject)wpJObject["webPartData"]["dynamicDataValues"];
                }

                if (wpJObject["webPartData"] != null && wpJObject["webPartData"]["id"] != null)
                {
                    this.webPartId = wpJObject["webPartData"]["id"].Value<string>();
                }

                this.usingSpControlDataOnly = true;
            }
            else
            {
                this.webPartData = wpDiv.GetAttribute(ClientSideWebPart.WebPartAttribute);

                // Decode the html encoded string
                var decoded = WebUtility.HtmlDecode(wpDiv.GetAttribute(ClientSideWebPart.WebPartDataAttribute));
                JObject wpJObject = JObject.Parse(decoded);
                this.title = wpJObject["title"] != null ? wpJObject["title"].Value<string>() : "";
                this.description = wpJObject["description"] != null ? wpJObject["description"].Value<string>() : "";
                // Set property to trigger correct loading of properties 
                this.PropertiesJson = wpJObject["properties"].ToString(Formatting.None);

                // Set/update dataVersion if it was set in the json data
                if (!string.IsNullOrEmpty(wpJObject["dataVersion"].ToString(Formatting.None)))
                {
                    this.dataVersion = wpJObject["dataVersion"].ToString(Formatting.None).Trim('"');
                }

                // Check for fullbleed supporting web parts
                if (wpJObject["properties"] != null && wpJObject["properties"]["isFullWidth"] != null)
                {
                    this.supportsFullBleed = wpJObject["properties"]["isFullWidth"].Value<Boolean>();
                }

                // Store the server processed content as that's needed for full fidelity
                if (wpJObject["serverProcessedContent"] != null && wpJObject["serverProcessedContent"].ToString() != "")
                {
                    this.serverProcessedContent = (JObject)wpJObject["serverProcessedContent"];
                }

                if (wpJObject["dynamicDataPaths"] != null)
                {
                    this.dynamicDataPaths = (JObject)wpJObject["dynamicDataPaths"];
                }

                if (wpJObject["dynamicDataValues"] != null)
                {
                    this.dynamicDataValues = (JObject)wpJObject["dynamicDataValues"];
                }

                this.webPartId = wpJObject["id"].Value<string>();

                var wpHtmlProperties = wpDiv.GetElementsByTagName("div").Where(a => a.HasAttribute(ClientSideWebPart.WebPartHtmlPropertiesAttribute)).FirstOrDefault();
                this.htmlPropertiesData = wpHtmlProperties.InnerHtml;
                this.htmlProperties = wpHtmlProperties.GetAttribute(ClientSideWebPart.WebPartHtmlPropertiesAttribute);
            }
        }

        private void SetPropertiesJson(string json)
        {
            if (String.IsNullOrEmpty(json))
            {
                json = "{}";
            }

            this.propertiesJson = json;

            var parsedJson = JObject.Parse(json);

            // If the passed structure is the top level JSON structure, which it typically is, then grab the properties from it
            if (parsedJson["webPartData"] != null && parsedJson["webPartData"]["properties"] != null)
            {
                this.properties = (JObject)parsedJson["webPartData"]["properties"];
            }
            else if (parsedJson["properties"] != null)
            {
                this.properties = (JObject)parsedJson["properties"];
            }
            else
            {
                this.properties = parsedJson;
            }

            // Get the web part data version if supplied by the web part json properties
            if (parsedJson["webPartData"] != null && parsedJson["webPartData"]["dataVersion"] != null)
            {
                this.dataVersion = parsedJson["webPartData"]["dataVersion"].ToString(Formatting.None).Trim('"');

            }
            else if (parsedJson["dataVersion"] != null)
            {
                this.dataVersion = parsedJson["dataVersion"].ToString(Formatting.None).Trim('"');
            }

            // If the web part has the serverProcessedContent property then keep this one as it might be needed as input to render the web part HTML later on
            if (parsedJson["webPartData"] != null && parsedJson["webPartData"]["serverProcessedContent"] != null)
            {
                this.serverProcessedContent = (JObject)parsedJson["webPartData"]["serverProcessedContent"];
            }
            else if (parsedJson["serverProcessedContent"] != null)
            {
                this.serverProcessedContent = (JObject)parsedJson["serverProcessedContent"];
            }

            if (parsedJson["webPartData"] != null && parsedJson["webPartData"]["dynamicDataPaths"] != null)
            {
                this.dynamicDataPaths = (JObject)parsedJson["webPartData"]["dynamicDataPaths"];
            }
            else if (parsedJson["dynamicDataPaths"] != null)
            {
                this.dynamicDataPaths = (JObject)parsedJson["dynamicDataPaths"];
            }

            if (parsedJson["webPartData"] != null && parsedJson["webPartData"]["dynamicDataValues"] != null)
            {
                this.dynamicDataValues = (JObject)parsedJson["webPartData"]["dynamicDataValues"];
            }
            else if (parsedJson["dynamicDataValues"] != null)
            {
                this.dynamicDataValues = (JObject)parsedJson["dynamicDataValues"];
            }
        }
            #endregion
    }
#endif
        }
