using AngleSharp.Dom;
using AngleSharp.Extensions;
using AngleSharp.Parser.Html;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Text;
#if !NETSTANDARD2_0
using System.Web.UI;
#endif

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Controls of type 4 ( = text control)
    /// </summary>
    public class ClientSideText : CanvasControl
    {
        #region variables
        public const string TextRteAttribute = "data-sp-rte";

        private string rte;
        private ClientSideTextControlData spControlData;
        private string previewText;
        #endregion

        #region construction
        /// <summary>
        /// Creates a <see cref="ClientSideText"/> instance
        /// </summary>
        public ClientSideText() : base()
        {
            this.controlType = 4;
            this.rte = "";
            this.previewText = "";
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
        /// Text used in page preview in news web part
        /// </summary>
        public string PreviewText
        {
            get
            {
                return this.previewText;
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

            try
            {
                var nodeList = new HtmlParser().ParseFragment(this.Text, null);
                this.previewText = string.Concat(nodeList.Select(x => x.Text()));
            }
            catch { }

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
#endif
}
