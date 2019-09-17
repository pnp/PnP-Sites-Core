using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a CanvasSection
    /// </summary>
    public partial class CanvasSection : BaseModel, IEquatable<CanvasSection>
    {
        #region Private Members

        private CanvasControlCollection _controls;

        #endregion

        #region Public Members

        /// <summary>
        /// Gets or sets the controls
        /// </summary>
        public CanvasControlCollection Controls
        {
            get { return _controls; }
            private set { _controls = value; }
        }

        /// <summary>
        /// Defines the order of the Canvas section for a Client-side Page.
        /// </summary>
        public float Order { get; set; }

        /// <summary>
        /// Defines the type of the Canvas section for a Client-side Page.
        /// </summary>
        public CanvasSectionType Type { get; set; }

        /// <summary>
        /// The emphasis color of the Canvas Section for a Client-side Page
        /// </summary>
        public Emphasis BackgroundEmphasis { get; set; }

        /// <summary>
        /// The emphasis color of the Canvas Section for a Client-side Page
        /// </summary>
        public Emphasis VerticalSectionEmphasis { get; set; }
        
        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for CanvasSection class
        /// </summary>
        public CanvasSection()
        {
            this._controls = new CanvasControlCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|",
                this.Controls.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                Order.GetHashCode(),
                Type.GetHashCode(),
                BackgroundEmphasis.GetHashCode(),
                VerticalSectionEmphasis.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with CanvasSection class
        /// </summary>
        /// <param name="obj">Object that represents CanvasSection</param>
        /// <returns>Checks whether object is CanvasSection class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is CanvasSection))
            {
                return (false);
            }
            return (Equals((CanvasSection)obj));
        }

        /// <summary>
        /// Compares CanvasSection object based on Controls, Order, Type, BackgroundEmphasis, and VerticalSectionEmphasis
        /// </summary>
        /// <param name="other">CanvasSection Class object</param>
        /// <returns>true if the CanvasSection object is equal to the current object; otherwise, false.</returns>
        public bool Equals(CanvasSection other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Controls.DeepEquals(other.Controls) &&
                this.Order == other.Order &&
                this.Type == other.Type &&
                this.BackgroundEmphasis == other.BackgroundEmphasis &&
                this.VerticalSectionEmphasis == other.VerticalSectionEmphasis
                );
        }

        #endregion
    }

    /// <summary>
    /// The type of the Canvas section for a Client-side Page.
    /// </summary>
    public enum CanvasSectionType
    {
        /// <summary>
        /// One column
        /// </summary>
        OneColumn,
        /// <summary>
        /// One column, full browser width. This one only works for communication sites in combination with image or hero webparts
        /// </summary>
        OneColumnFullWidth,
        /// <summary>
        /// Two columns of the same size
        /// </summary>
        TwoColumn,
        /// <summary>
        /// Three columns of the same size
        /// </summary>
        ThreeColumn,
        /// <summary>
        /// Two columns, left one is 2/3, right one 1/3
        /// </summary>
        TwoColumnLeft,
        /// <summary>
        /// Two columns, left one is 1/3, right one 2/3
        /// </summary>
        TwoColumnRight,
        /// <summary>
        /// One column, and a vertical section
        /// </summary>
        OneColumnVerticalSection,
        /// <summary>
        /// Two columns of the same size, and a vertical section
        /// </summary>
        TwoColumnVerticalSection,
        /// <summary>
        /// Three columns of the same size, and a vertical section
        /// </summary>
        ThreeColumnVerticalSection,
        /// <summary>
        /// Two columns, left one is 2/3, right one 1/3, and a vertical section
        /// </summary>
        TwoColumnLeftVerticalSection,
        /// <summary>
        /// Two columns, left one is 1/3, right one 2/3, and a vertical section
        /// </summary>
        TwoColumnRightVerticalSection,
    }
}
