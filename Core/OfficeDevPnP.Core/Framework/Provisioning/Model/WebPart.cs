using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class WebPart : BaseModel, IEquatable<WebPart>
    {
        #region Properties
        /// <summary>
        /// Webpart Row
        /// </summary>
        public uint Row { get; set; }
        /// <summary>
        /// Webpart Column
        /// </summary>
        public uint Column { get; set; }
        /// <summary>
        /// Webpart Title
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// Webpart Contents
        /// </summary>
        public string Contents { get; set; }
        /// <summary>
        /// Webpart Zone
        /// </summary>
        public string Zone { get; set; }
        /// <summary>
        /// Webpart Order
        /// </summary>
        public uint Order { get; set; }
        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}",
                this.Row.GetHashCode(),
                this.Column.GetHashCode(),
                (this.Contents != null ? this.Contents.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with WebPart
        /// </summary>
        /// <param name="obj">Object that represents WebPart</param>
        /// <returns>true if the current object is equal to the WebPart</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is File))
            {
                return (false);
            }
            return (Equals((File)obj));
        }

        /// <summary>
        /// Compares WebPart object based on Row, Column and Contents
        /// </summary>
        /// <param name="other">WebPart object</param>
        /// <returns>true if the WebPart object is equal to the current object; otherwise, false.</returns>
        public bool Equals(WebPart other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Row == other.Row &&
                this.Column == other.Column &&
                this.Contents == other.Contents);
        }

        #endregion
    }
}
