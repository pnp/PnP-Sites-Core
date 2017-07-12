using AngleSharp.Dom;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.UI;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    #region Client Side control classes
    /// <summary>
    /// Base class for a canvas control 
    /// </summary>
    public abstract class CanvasControl
    {
        #region variables
        public const string CanvasControlAttribute = "data-sp-canvascontrol";
        public const string CanvasDataVersionAttribute = "data-sp-canvasdataversion";
        public const string ControlDataAttribute = "data-sp-controldata";

        internal int order;
        internal int controlType;
        internal string jsonControlData;
        internal string dataVersion;
        internal string canvasControlData;
        internal Guid instanceId;
        internal CanvasZone zone;
        internal CanvasSection section;
        #endregion

        #region construction
        /// <summary>
        /// Constructs the canvas control
        /// </summary>
        public CanvasControl()
        {
            this.dataVersion = "1.0";
            this.instanceId = Guid.NewGuid();
            this.canvasControlData = "";
            this.order = 0;
        }
        #endregion

        #region Properties
        /// <summary>
        /// The <see cref="CanvasZone"/> hosting this control
        /// </summary>
        public CanvasZone Zone
        {
            get
            {
                return this.zone;
            }
        }

        /// <summary>
        /// The <see cref="CanvasSection"/> hosting this control
        /// </summary>
        public CanvasSection Section
        {
            get
            {
                return this.section;
            }
        }

        /// <summary>
        /// The internal storage version used for this control
        /// </summary>
        public string DataVersion
        {
            get
            {
                return dataVersion;
            }
        }

        /// <summary>
        /// Value of the control's "data-sp-canvascontrol" attribute
        /// </summary>
        public string CanvasControlData
        {
            get
            {
                return canvasControlData;
            }
        }

        /// <summary>
        /// Type of the control: 3 is a text part, 4 is a client side web part
        /// </summary>
        public int ControlType
        {
            get
            {
                return controlType;
            }
        }

        /// <summary>
        /// Value of the control's "data-sp-controldata" attribute
        /// </summary>
        public string JsonControlData
        {
            get
            {
                return jsonControlData;
            }
        }

        /// <summary>
        /// Instance ID of the control
        /// </summary>
        public Guid InstanceId
        {
            get
            {
                return instanceId;
            }
        }

        /// <summary>
        /// Order of the control in the control collection
        /// </summary>
        public int Order
        {
            get
            {
                return this.order;
            }
            set
            {
                this.order = value;
            }
        }

        /// <summary>
        /// Type if the control (<see cref="ClientSideText"/> or <see cref="ClientSideWebPart"/>)
        /// </summary>
        public abstract Type Type { get; }
        #endregion

        #region public methods
        /// <summary>
        /// Converts a control object to it's html representation
        /// </summary>
        /// <param name="controlIndex">The sequence of the control inside the section</param>
        /// <returns>Html representation of a control</returns>
        public abstract string ToHtml(int controlIndex);

        /// <summary>
        /// Removes the control from the page
        /// </summary>
        public void Delete()
        {
            this.Section.Zone.Page.Controls.Remove(this);
        }

        /// <summary>
        /// Receives "data-sp-controldata" content and detects the type of the control
        /// </summary>
        /// <param name="controlDataJson">data-sp-controldata json string</param>
        /// <returns>Type of the control represented by the json string</returns>
        public static Type GetType(string controlDataJson)
        {
            if (controlDataJson == null)
            {
                throw new ArgumentNullException("ControlDataJson cannot be null");
            }

            // Decode the html encoded string
            var decoded = WebUtility.HtmlDecode(controlDataJson);

            // Deserialize the json string
            var jsonSerializerSettings = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            var controlData = JsonConvert.DeserializeObject<ClientSideCanvasControlData>(decoded, jsonSerializerSettings);

            if (controlData.ControlType == 3)
            {
                return typeof(ClientSideWebPart);
            }
            else if (controlData.ControlType == 4)
            {
                return typeof(ClientSideText);
            }
            else if (controlData.ControlType == 0)
            {
                return typeof(CanvasSection);
            }

            return null;
        }
        #endregion

        #region Internal and private methods
        internal virtual void FromHtml(IElement element)
        {
            // deserialize control data
            var jsonSerializerSettings = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore
            };

            var controlData = JsonConvert.DeserializeObject<ClientSideCanvasControlData>(element.GetAttribute(CanvasControl.ControlDataAttribute), jsonSerializerSettings);

            // populate base object
            this.dataVersion = element.GetAttribute(CanvasControl.CanvasDataVersionAttribute);
            this.canvasControlData = element.GetAttribute(CanvasControl.CanvasControlAttribute);
            this.controlType = controlData.ControlType;
            this.instanceId = new Guid(controlData.Id);
        }

        #endregion
    }

    /// <summary>
    /// Controls of type 4 ( = text control)
    /// </summary>
    public class ClientSideText : CanvasControl
    {
        #region variables
        public const string TextRteAttribute = "data-sp-rte";

        private string rte;
        private ClientSideTextControlData spControlData;
        #endregion

        #region construction
        /// <summary>
        /// Creates a <see cref="ClientSideText"/> instance
        /// </summary>
        public ClientSideText() : base()
        {
            this.controlType = 4;
            this.rte = "";
        }
        #endregion

        #region Properties
        /// <summary>
        /// Text value of the client side text control
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Value of the "data-sp-rte" attribute
        /// </summary>
        public string Rte
        {
            get
            {
                return this.rte;
            }
        }

        /// <summary>
        /// Type of the control (= <see cref="ClientSideText"/>)
        /// </summary>
        public override Type Type
        {
            get
            {
                return typeof(ClientSideText);
            }
        }

        /// <summary>
        /// Deserialized value of the "data-sp-controldata" attribute
        /// </summary>
        public ClientSideTextControlData SpControlData
        {
            get
            {
                return this.spControlData;
            }
        }
        #endregion

        #region public methods
        /// <summary>
        /// Converts this <see cref="ClientSideText"/> control to it's html representation
        /// </summary>
        /// <param name="controlIndex">The sequence of the control inside the section</param>
        /// <returns>Html representation of this <see cref="ClientSideText"/> control</returns>
        public override string ToHtml(int controlIndex)
        {
            // Can this control be hosted in this zone type?
            if (this.Zone.Type == CanvasZoneTemplate.OneColumnFullWidth)
            {
                throw new Exception("You cannot host text controls inside a one column full width zone, only an image web part or hero web part are allowed");
            }

            // Obtain the json data
            ClientSideTextControlData controlData = new ClientSideTextControlData() {
                ControlType = this.ControlType,
                Id = this.InstanceId.ToString("D"),
                Position = new ClientSideCanvasControlPosition()
                {
                    ZoneIndex = this.Zone.Order,
                    SectionIndex = this.Section.Order,
                    SectionFactor = this.Section.SectionFactor,
                    ControlIndex = controlIndex,
                },
                EditorType = "CKEditor" };
            jsonControlData = JsonConvert.SerializeObject(controlData);

            StringBuilder html = new StringBuilder(100);
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;

                htmlWriter.AddAttribute(CanvasControlAttribute, this.CanvasControlData);
                htmlWriter.AddAttribute(CanvasDataVersionAttribute, this.DataVersion);
                htmlWriter.AddAttribute(ControlDataAttribute, this.JsonControlData);
                htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);

                htmlWriter.AddAttribute(TextRteAttribute, this.Rte);
                htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);

                htmlWriter.RenderBeginTag(HtmlTextWriterTag.P);
                htmlWriter.Write(this.Text);
                htmlWriter.RenderEndTag();

                htmlWriter.RenderEndTag();
                htmlWriter.RenderEndTag();
            }

            return html.ToString();
        }
        #endregion

        #region Internal and private methods
        internal override void FromHtml(IElement element)
        {
            base.FromHtml(element);

            var div = element.GetElementsByTagName("div").Where(a => a.HasAttribute(TextRteAttribute)).FirstOrDefault();
            this.rte = div.GetAttribute(TextRteAttribute);
            this.Text = div.InnerHtml;

            // load data from the data-sp-controldata attribute
            var jsonSerializerSettings = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            this.spControlData = JsonConvert.DeserializeObject<ClientSideTextControlData>(element.GetAttribute(CanvasControl.ControlDataAttribute), jsonSerializerSettings);
            this.controlType = this.spControlData.ControlType;
        }
        #endregion
    }

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
        private string description;
        private string propertiesJson;
        private ClientSideWebPartControlData spControlData;
        private JObject properties;
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
            this.SetPropertiesJson("{}");
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
        /// Json serialized web part properties
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
        public override string ToHtml(int controlIndex)
        {
            // Can this control be hosted in this zone type?
            if (this.Zone.Type == CanvasZoneTemplate.OneColumnFullWidth)
            {
                if (!this.WebPartId.Equals(ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Image), StringComparison.InvariantCultureIgnoreCase) ||
                    !this.WebPartId.Equals(ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Hero), StringComparison.InvariantCultureIgnoreCase))
                {
                    throw new Exception("You cannot host this web part inside a one column full width zone, only an image web part or hero web part are allowed");
                }
            }

            // Obtain the json data
            ClientSideWebPartControlData controlData = new ClientSideWebPartControlData()
            {
                ControlType = this.ControlType,
                Id = this.InstanceId.ToString("D"),
                WebPartId = this.WebPartId,
                Position = new ClientSideCanvasControlPosition()
                {
                    ZoneIndex = this.Zone.Order,
                    SectionIndex = this.Section.Order,
                    SectionFactor = this.Section.SectionFactor,
                    ControlIndex = controlIndex,
                },
            };
            ClientSideWebPartData webpartData = new ClientSideWebPartData() { Id = controlData.WebPartId, InstanceId = controlData.Id, Title = this.Title, Description = this.Description, DataVersion = this.DataVersion, Properties = "jsonPropsToReplacePnPRules" };

            this.jsonControlData = JsonConvert.SerializeObject(controlData);
            this.jsonWebPartData = JsonConvert.SerializeObject(webpartData);
            this.jsonWebPartData = jsonWebPartData.Replace("\"jsonPropsToReplacePnPRules\"", this.Properties.ToString(Formatting.None));

            StringBuilder html = new StringBuilder(100);
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;
                htmlWriter.AddAttribute(CanvasControlAttribute, this.CanvasControlData);
                htmlWriter.AddAttribute(CanvasDataVersionAttribute, this.DataVersion);
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
                htmlWriter.Write(this.HtmlPropertiesData);
                htmlWriter.RenderEndTag();

                htmlWriter.RenderEndTag();
                htmlWriter.RenderEndTag();
            }

            return html.ToString();
        }
        #endregion

        #region Internal and private methods
        internal override void FromHtml(IElement element)
        {
            base.FromHtml(element);

            // load data from the data-sp-controldata attribute
            var jsonSerializerSettings = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            this.spControlData = JsonConvert.DeserializeObject<ClientSideWebPartControlData>(element.GetAttribute(CanvasControl.ControlDataAttribute), jsonSerializerSettings);
            this.controlType = this.spControlData.ControlType;


            var wpDiv = element.GetElementsByTagName("div").Where(a => a.HasAttribute(ClientSideWebPart.WebPartDataAttribute)).FirstOrDefault();
            this.webPartData = wpDiv.GetAttribute(ClientSideWebPart.WebPartAttribute);

            // Decode the html encoded string
            var decoded = WebUtility.HtmlDecode(wpDiv.GetAttribute(ClientSideWebPart.WebPartDataAttribute));
            JObject wpJObject = JObject.Parse(decoded);
            this.title = wpJObject["title"] != null ? wpJObject["title"].Value<string>() : "";
            this.description = wpJObject["description"] != null ? wpJObject["description"].Value<string>() : "";
            this.propertiesJson = wpJObject["properties"].ToString(Formatting.None);
            this.webPartId = wpJObject["id"].Value<string>();

            var wpHtmlProperties = wpDiv.GetElementsByTagName("div").Where(a => a.HasAttribute(ClientSideWebPart.WebPartHtmlPropertiesAttribute)).FirstOrDefault();
            this.htmlPropertiesData = wpHtmlProperties.InnerHtml;
            this.htmlProperties = wpHtmlProperties.GetAttribute(ClientSideWebPart.WebPartHtmlPropertiesAttribute);
        }

        private void SetPropertiesJson(string json)
        {
            if (String.IsNullOrEmpty(json))
            {
                json = "{}";
            }

            this.propertiesJson = json;
            this.properties = JObject.Parse(json);
        }
        #endregion
    }
    #endregion
#endif
}
