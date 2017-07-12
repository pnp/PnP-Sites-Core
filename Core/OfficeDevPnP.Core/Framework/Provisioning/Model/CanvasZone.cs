using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a CanvasZone
    /// </summary>
    public partial class CanvasZone : BaseModel, IEquatable<CanvasZone>
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
        /// Defines the order of the Canvas Zone for a Client-side Page.
        /// </summary>
        public Int32 Order { get; set; }

        /// <summary>
        /// Defines the type of the Canvas Zone for a Client-side Page.
        /// </summary>
        public CanvasZoneType Type { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for CanvasZone class
        /// </summary>
        public CanvasZone()
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
            return (String.Format("{0}|{1}|{2}|",
                this.Controls.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                Order.GetHashCode(),
                Type.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with CanvasZone class
        /// </summary>
        /// <param name="obj">Object that represents CanvasZone</param>
        /// <returns>Checks whether object is CanvasZone class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is CanvasZone))
            {
                return (false);
            }
            return (Equals((CanvasZone)obj));
        }

        /// <summary>
        /// Compares CanvasZone object based on Controls, Order, and Type
        /// </summary>
        /// <param name="other">CanvasZone Class object</param>
        /// <returns>true if the CanvasZone object is equal to the current object; otherwise, false.</returns>
        public bool Equals(CanvasZone other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Controls.DeepEquals(other.Controls) &&
                this.Order == other.Order &&
                this.Type == other.Type
                );
        }

        #endregion
    }

    /// <summary>
    /// The type of the Canvas Zone for a Client-side Page.
    /// </summary>
    public enum CanvasZoneType
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
    }
}
