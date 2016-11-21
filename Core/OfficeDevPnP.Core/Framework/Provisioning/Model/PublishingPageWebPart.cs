using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class PublishingPageWebPart : WebPart
    {
        #region Properties

        public string DefaultViewDisplayName { get; set; }
        public bool IsListViewWebPart
        {
            get
            {
                return DefaultViewDisplayName != null;
            }
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return String.Format("{0}|{1}|{2}|{3}",
                base.GetHashCode(),
                DefaultViewDisplayName.GetHashCode()).GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (!(obj is PublishingPageWebPart))
            {
                return (false);
            }
            return (Equals((PublishingPageWebPart)obj));
        }

        public bool Equals(PublishingPageWebPart other)
        {
            if (other == null)
            {
                return (false);
            }

            return (base.Equals(other) &&
                    this.DefaultViewDisplayName == other.DefaultViewDisplayName);
        }

        #endregion
    }
}
