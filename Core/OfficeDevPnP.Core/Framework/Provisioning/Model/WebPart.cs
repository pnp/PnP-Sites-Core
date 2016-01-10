using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class WebPart : BaseModel, IEquatable<WebPart>
    {
        #region Properties
        public uint Row { get; set; }

        public uint Column { get; set; }

        public string Title { get; set; }

        public string Contents { get; set; }

        public string Zone { get; set; }

        public uint Order { get; set; }
        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}",
                this.Row.GetHashCode(),
                this.Column.GetHashCode(),
                (this.Contents != null ? this.Contents.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is File))
            {
                return (false);
            }
            return (Equals((File)obj));
        }

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
