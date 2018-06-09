using AngleSharp.Dom;
using Newtonsoft.Json;
using System;
using System.Linq;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
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
        /// The <see cref="CanvasSection"/> hosting  this control
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
        /// Type of the control: 4 is a text part, 3 is a client side web part
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
#endif
}
