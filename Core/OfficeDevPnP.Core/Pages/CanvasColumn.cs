using Newtonsoft.Json;
using System;
using System.Linq;
using System.Text;
#if !NETSTANDARD2_0
using System.Web.UI;
#endif

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016

    /// <summary>
    /// Represents a column in a canvas section
    /// </summary>
    public class CanvasColumn
    {
        #region variables
        public const string CanvasControlAttribute = "data-sp-canvascontrol";
        public const string CanvasDataVersionAttribute = "data-sp-canvasdataversion";
        public const string ControlDataAttribute = "data-sp-controldata";

        private int columnFactor;
        private int layoutIndex;
        private int? zoneEmphasis;
        private CanvasSection section;
        private string DataVersion = "1.0";
        #endregion

        // internal constructors as we don't want users to manually create sections
        #region construction
        internal CanvasColumn(CanvasSection section)
        {
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            this.section = section;
            this.columnFactor = 12;
            this.Order = 0;
            this.layoutIndex = 1;
        }

        internal CanvasColumn(CanvasSection section, int order)
        {
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            this.section = section;
            this.Order = order;
            this.layoutIndex = 1;
        }

        internal CanvasColumn(CanvasSection section, int order, int? sectionFactor)
        {
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            this.section = section;
            this.Order = order;
            this.columnFactor = sectionFactor.HasValue ? sectionFactor.Value : 12;
            this.layoutIndex = 1;
        }

        internal CanvasColumn(CanvasSection section, int order, int? sectionFactor, int? layoutIndex)
        {
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            this.section = section;
            this.Order = order;
            this.columnFactor = sectionFactor.HasValue ? sectionFactor.Value : 12;
            this.layoutIndex = layoutIndex.HasValue ? layoutIndex.Value : 1;
        }
        #endregion

        #region Properties
        internal int Order { get; set; }

        /// <summary>
        /// <see cref="CanvasSection"/> this section belongs to
        /// </summary>
        public CanvasSection Section
        {
            get
            {
                return this.section;
            }
        }

        /// <summary>
        /// Column size factor. Max value is 12 (= one column), other options are 8,6,4 or 0
        /// </summary>
        public int ColumnFactor
        {
            get
            {
                return this.columnFactor;
            }
        }

        /// <summary>
        /// Returns the layout index. Defaults to 1, except for the vertical section column this is 2
        /// </summary>
        public int LayoutIndex
        {
            get
            {
                return this.layoutIndex;
            }
        }

        /// <summary>
        /// List of <see cref="CanvasControl"/> instances that are hosted in this section
        /// </summary>
        public System.Collections.Generic.List<CanvasControl> Controls
        {
            get
            {
                return this.Section.Page.Controls.Where(p => p.Section == this.Section && p.Column == this).ToList<CanvasControl>();
            }
        }

        /// <summary>
        /// Is this a vertical section column?
        /// </summary>
        public bool IsVerticalSectionColumn
        {
            get
            {
                return this.LayoutIndex == 2;
            }
        }

        /// <summary>
        /// Color emphasis of the column (used for the vertical section column) 
        /// </summary>
        public int? VerticalSectionEmphasis
        {
            get
            {
                if (this.LayoutIndex == 2)
                {
                    return this.zoneEmphasis;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (this.LayoutIndex == 2)
                {
                    if (value < 0 || value > 3)
                    {
                        throw new ArgumentException($"The zoneEmphasis value needs to be between 0 and 3. See the Microsoft.SharePoint.Client.SPVariantThemeType values for the why.");
                    }

                    this.zoneEmphasis = value;
                }
            }
        }
        #endregion

        #region public methods
        /// <summary>
        /// Renders a HTML presentation of this section
        /// </summary>
        /// <returns>The HTML presentation of this section</returns>
        public string ToHtml()
        {
            StringBuilder html = new StringBuilder(100);
#if !NETSTANDARD2_0
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;
#endif
                bool controlWrittenToSection = false;
                int controlIndex = 0;
                foreach (var control in this.Section.Page.Controls.Where(p => p.Section == this.Section && p.Column == this).OrderBy(z => z.Order))
                {
                    controlIndex++;
#if NETSTANDARD2_0
                    html.Append(control.ToHtml(controlIndex));
#else
                    htmlWriter.Write(control.ToHtml(controlIndex));
#endif
                    controlWrittenToSection = true;
                }

                // if a section does not contain a control we still need to render it, otherwise it get's "lost"
                if (!controlWrittenToSection)
                {
                    // Obtain the json data
                    var clientSideCanvasPosition = new ClientSideCanvasData()
                    {
                        Position = new ClientSideCanvasPosition()
                        {
                            ZoneIndex = this.Section.Order,
                            SectionIndex = this.Order,
                            SectionFactor = this.ColumnFactor,
#if !SP2019
                            LayoutIndex = this.LayoutIndex,
#endif
                        },

                        Emphasis = new ClientSideSectionEmphasis()
                        {
                            ZoneEmphasis = this.Section.ZoneEmphasis,
                        }
                    };

                    var jsonControlData = JsonConvert.SerializeObject(clientSideCanvasPosition);

#if NETSTANDARD2_0
                html.Append($@"<div {CanvasControlAttribute}="""" {CanvasDataVersionAttribute}=""{this.DataVersion}"" {ControlDataAttribute}=""{jsonControlData.Replace("\"", "&quot;")}""></div>");
#else
                    htmlWriter.NewLine = string.Empty;

                    htmlWriter.AddAttribute(CanvasControlAttribute, "");
                    htmlWriter.AddAttribute(CanvasDataVersionAttribute, this.DataVersion);
                    htmlWriter.AddAttribute(ControlDataAttribute, jsonControlData);
                    htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);
                    htmlWriter.RenderEndTag();
#endif
                }
#if !NETSTANDARD2_0
            }
#endif

            return html.ToString();
        }

        /// <summary>
        /// Resets the column, used in scenarios where a section is changed from type (e.g. from 3 column to 2 column)
        /// </summary>
        /// <param name="order">Column order to set</param>
        /// <param name="columnFactor">Column factor to set</param>
        public void ResetColumn(int order, int columnFactor)
        {
            this.Order = order;
            this.columnFactor = columnFactor;
        }

        #region Internal and helper methods
        internal void MoveTo(CanvasSection section)
        {
            this.section = section;
        }
        #endregion

        #endregion
    }
#endif
}
