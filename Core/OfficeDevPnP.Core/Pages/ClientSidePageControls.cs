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
        internal CanvasSection section;
        internal CanvasColumn column;
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
        /// The <see cref="CanvasColumn"/> hosting this control
        /// </summary>
        public CanvasColumn Column
        {
            get
            {
                return this.column;
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
        public abstract string ToHtml(float controlIndex);

        /// <summary>
        /// Removes the control from the page
        /// </summary>
        public void Delete()
        {
            this.Column.Section.Page.Controls.Remove(this);
        }

        /// <summary>
        /// Moves the control to another section and column
        /// </summary>
        /// <param name="newSection">New section that will host the control</param>
        public void Move(CanvasSection newSection)
        {
            this.section = newSection;
            this.column = newSection.DefaultColumn;
        }

        /// <summary>
        /// Moves the control to another section and column
        /// </summary>
        /// <param name="newSection">New section that will host the control</param>
        /// <param name="order">New order for the control in the new section</param>
        public void Move(CanvasSection newSection, int order)
        {
            Move(newSection);
            this.order = order;
        }

        /// <summary>
        /// Moves the control to another section and column
        /// </summary>
        /// <param name="newColumn">New column that will host the control</param>
        public void Move(CanvasColumn newColumn)
        {
            this.section = newColumn.Section;
            this.column = newColumn;
        }

        /// <summary>
        /// Moves the control to another section and column
        /// </summary>
        /// <param name="newColumn">New column that will host the control</param>
        /// <param name="order">New order for the control in the new column</param>
        public void Move(CanvasColumn newColumn, int order)
        {
            Move(newColumn);
            this.order = order;
        }

        /// <summary>
        /// Moves the control to another section and column while keeping it's current position
        /// </summary>
        /// <param name="newSection">New section that will host the control</param>
        public void MovePosition(CanvasSection newSection)
        {
            var currentSection = this.Section;
            this.section = newSection;
            this.column = newSection.DefaultColumn;
            ReindexSection(currentSection);
            ReindexSection(this.Section);
        }

        /// <summary>
        /// Moves the control to another section and column in the given position
        /// </summary>
        /// <param name="newSection">New section that will host the control</param>
        /// <param name="position">New position for the control in the new section</param>
        public void MovePosition(CanvasSection newSection, int position)
        {
            var currentSection = this.Section;
            MovePosition(newSection);
            ReindexSection(currentSection);
            MovePosition(position);
        }

        /// <summary>
        /// Moves the control to another section and column while keeping it's current position
        /// </summary>
        /// <param name="newColumn">New column that will host the control</param>
        public void MovePosition(CanvasColumn newColumn)
        {
            var currentColumn = this.Column;
            this.section = newColumn.Section;
            this.column = newColumn;
            ReindexColumn(currentColumn);
            ReindexColumn(this.Column);
        }

        /// <summary>
        /// Moves the control to another section and column in the given position
        /// </summary>
        /// <param name="newColumn">New column that will host the control</param>
        /// <param name="position">New position for the control in the new column</param>
        public void MovePosition(CanvasColumn newColumn, int position)
        {
            var currentColumn = this.Column;
            MovePosition(newColumn);
            ReindexColumn(currentColumn);
            MovePosition(position);
        }

        /// <summary>
        /// Moves the control inside the current column to a new position
        /// </summary>
        /// <param name="position">New position for this control</param>
        public void MovePosition(int position)
        {
            // Ensure we're having a clean sequence before starting
            ReindexColumn();

            if (position > this.Order)
            {
                position++;
            }

            foreach (var control in this.section.Page.Controls.Where(c => c.Section == this.section && c.Column == this.column && c.Order >= position).OrderBy(p => p.Order))
            {
                control.Order = control.Order + 1;
            }
            this.Order = position;

            // Ensure we're having a clean sequence to return
            ReindexColumn();
        }

        private void ReindexColumn()
        {
            ReindexColumn(this.Column);
        }

        private void ReindexColumn(CanvasColumn column)
        {
            var index = 0;
            foreach (var control in this.column.Section.Page.Controls.Where(c => c.Section == column.Section && c.Column == column).OrderBy(c => c.Order))
            {
                index++;
                control.order = index;
            }
        }

        private void ReindexSection(CanvasSection section)
        {
            foreach (var column in section.Columns)
            {
                ReindexColumn(column);
            }
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

            // Deserialize the json string
            var jsonSerializerSettings = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            var controlData = JsonConvert.DeserializeObject<ClientSideCanvasControlData>(controlDataJson, jsonSerializerSettings);

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
                return typeof(CanvasColumn);
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
        public override string ToHtml(float controlIndex)
        {
            // Can this control be hosted in this section type?
            if (this.Section.Type == CanvasSectionTemplate.OneColumnFullWidth)
            {
                throw new Exception("You cannot host text controls inside a one column full width section, only an image web part or hero web part are allowed");
            }

            // Obtain the json data
            ClientSideTextControlData controlData = new ClientSideTextControlData()
            {
                ControlType = this.ControlType,
                Id = this.InstanceId.ToString("D"),
                Position = new ClientSideCanvasControlPosition()
                {
                    ZoneIndex = this.Section.Order,
                    SectionIndex = this.Column.Order,
                    SectionFactor = this.Column.ColumnFactor,
                    ControlIndex = controlIndex,
                },
                EditorType = "CKEditor"
            };
            jsonControlData = JsonConvert.SerializeObject(controlData);

            StringBuilder html = new StringBuilder(100);
#if NETSTANDARD2_0
            html.Append($@"<div {CanvasControlAttribute}=""{this.CanvasControlData}"" {CanvasDataVersionAttribute}=""{ this.DataVersion}""  {ControlDataAttribute}=""{this.jsonControlData.Replace("\"", "&quot;")}"">");
            html.Append($@"<div {TextRteAttribute}=""{this.Rte}"">");
            html.Append($@"<p>{this.Text}</p>");
            html.Append("</div>");
            html.Append("</div>");
#else
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
#endif
            return html.ToString();
        }
        #endregion

        #region Internal and private methods
        internal override void FromHtml(IElement element)
        {
            base.FromHtml(element);

            var div = element.GetElementsByTagName("div").Where(a => a.HasAttribute(TextRteAttribute)).FirstOrDefault();

            if (div != null)
            {
                this.rte = div.GetAttribute(TextRteAttribute);
            }
            else
            {
                // supporting updated rendering of Text controls, no nested DIV tag with the data-sp-rte attribute...so HTML content is embedded at the root
                this.rte = "";
                div = element;
            }

            // By default simple plain text is wrapped in a Paragraph, need to drop it to avoid getting multiple paragraphs on page edits.
            // Only drop the paragraph tag when there's only one Paragraph element underneath the DIV tag
            if ((div.FirstChild != null && (div.FirstChild as IElement).TagName.Equals("P", StringComparison.InvariantCultureIgnoreCase)) &&
                (div.ChildElementCount == 1))
            {
                this.Text = (div.FirstChild as IElement).InnerHtml;
            }
            else
            {
                this.Text = div.InnerHtml;
            }

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
        private JObject serverProcessedContent;
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
        public override string ToHtml(float controlIndex)
        {
            // Can this control be hosted in this section type?
            if (this.Section.Type == CanvasSectionTemplate.OneColumnFullWidth)
            {
                if (!(this.WebPartId.Equals(ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Image), StringComparison.InvariantCultureIgnoreCase) ||
                      this.WebPartId.Equals(ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Hero), StringComparison.InvariantCultureIgnoreCase)))
                {
                    throw new Exception("You cannot host this web part inside a one column full width section, only an image web part or hero web part are allowed");
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
                    ZoneIndex = this.Section.Order,
                    SectionIndex = this.Column.Order,
                    SectionFactor = this.Column.ColumnFactor,
                    ControlIndex = controlIndex,
                },
            };
            ClientSideWebPartData webpartData = new ClientSideWebPartData() { Id = controlData.WebPartId, InstanceId = controlData.Id, Title = this.Title, Description = this.Description, DataVersion = this.DataVersion, Properties = "jsonPropsToReplacePnPRules" };

            this.jsonControlData = JsonConvert.SerializeObject(controlData);
            this.jsonWebPartData = JsonConvert.SerializeObject(webpartData);
            this.jsonWebPartData = jsonWebPartData.Replace("\"jsonPropsToReplacePnPRules\"", this.Properties.ToString(Formatting.None));

            StringBuilder html = new StringBuilder(100);
#if NETSTANDARD2_0
            html.Append($@"<div {CanvasControlAttribute}=""{this.CanvasControlData}"" {CanvasDataVersionAttribute}=""{this.DataVersion}"" {ControlDataAttribute}=""{this.JsonControlData.Replace("\"", "&quot;")}"">");
            html.Append($@"<div {WebPartAttribute}=""{this.WebPartData}"" {WebPartDataVersionAttribute}=""{this.DataVersion}"" {WebPartDataAttribute}=""{this.JsonWebPartData}"">");
            html.Append($@"<div {WebPartComponentIdAttribute}=""""");
            html.Append(this.WebPartId);
            html.Append("/div>");
            html.Append($@"<div {WebPartHtmlPropertiesAttribute}=""{this.HtmlProperties}"">");
            RenderHtmlProperties(ref html);
            html.Append("</div>");
            html.Append("</div>");
            html.Append("</div>");
#else
            var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), "");
            try
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
                // Allow for override of the HTML value rendering if this would be needed by controls
                RenderHtmlProperties(ref htmlWriter);
                htmlWriter.RenderEndTag();

                htmlWriter.RenderEndTag();
                htmlWriter.RenderEndTag();
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
            // Set property to trigger correct loading of properties 
            this.PropertiesJson = wpJObject["properties"].ToString(Formatting.None);
            // Store the server processed content as that's needed for full fidelity
            if (wpJObject["serverProcessedContent"] != null)
            {
                this.serverProcessedContent = (JObject)wpJObject["serverProcessedContent"];
            }

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

            // If the web part has the serverProcessedContent property then keep this one as it might be needed as input to render the web part HTML later on
            if (parsedJson["webPartData"] != null && parsedJson["webPartData"]["serverProcessedContent"] != null)
            {
                this.serverProcessedContent = (JObject)parsedJson["webPartData"]["serverProcessedContent"];
            }
            else if (parsedJson["serverProcessedContent"] != null)
            {
                this.serverProcessedContent = (JObject)parsedJson["serverProcessedContent"];
            }

        }
        #endregion
    }
    #endregion
#endif
}
