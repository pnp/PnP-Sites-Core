using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class PropertyBagEntry : BaseModel, IEquatable<PropertyBagEntry>
    {
        #region Properties

        public string Key { get; set; }

        public string Value { get; set; }

        public bool Indexed { get; set; }

        public bool Overwrite { get; set; }
        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                (this.Key != null ? this.Key.GetHashCode() : 0),
                (this.Value != null ? this.Value.GetHashCode() : 0),
                this.Indexed.GetHashCode(),
                this.Overwrite.GetHashCode()
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is PropertyBagEntry))
            {
                return (false);
            }
            return (Equals((PropertyBagEntry)obj));
        }

        public bool Equals(PropertyBagEntry other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Key == other.Key &&
                this.Value == other.Value &&
                this.Indexed == other.Indexed &&
                this.Overwrite == other.Overwrite);
        }

        #endregion
    }
}
