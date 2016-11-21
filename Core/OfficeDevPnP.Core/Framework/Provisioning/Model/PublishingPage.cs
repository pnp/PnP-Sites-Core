using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class PublishingPage : BaseModel
    {
        #region Properties

        public string FileName { get; set; }
        public string FullFileName { get { return FileName + ".aspx"; } }
        public string Title { get; set; }
        public string Layout { get; set; }
        public bool Overwrite { get; set; }
        public bool Publish { get; set; }
        public bool WelcomePage { get; set; }
        public string PublishingPageContent { get; set; }
        public List<PublishingPageWebPart> WebParts { get; set; } = new List<PublishingPageWebPart>();
        public Dictionary<string, string> Properties { get; set; } = new Dictionary<string, string>();

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}",
               FileName.GetHashCode(),
               Title.GetHashCode(),
               Layout.GetHashCode(),
               Overwrite.GetHashCode(),
               Publish.GetHashCode(),
               WelcomePage.GetHashCode(),
               PublishingPageContent != null ? PublishingPageContent.GetHashCode() : 0,
               WebParts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
               Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode())
           ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is PublishingPage))
            {
                return (false);
            }
            return (Equals((PublishingPage)obj));
        }

        public bool Equals(PublishingPage other)
        {
            if (other == null)
            {
                return (false);
            }

            return (FileName == other.FileName &&
                Title == other.Title &&
                Layout == other.Layout &&
                Overwrite == other.Overwrite &&
                Publish == other.Publish &&
                WelcomePage == other.WelcomePage &&
                PublishingPageContent == other.PublishingPageContent &&
                WebParts.DeepEquals(other.WebParts) &&
                Properties.DeepEquals(other.Properties));
        }

        #endregion
    }
}
