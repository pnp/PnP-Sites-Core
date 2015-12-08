using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class Page : IEquatable<Page>
    {
        #region Private Members

        private List<WebPart> _webParts = new List<WebPart>();
        private ObjectSecurity _security = null;

        #endregion

        #region Properties

        public string Url { get; set; }

        public WikiPageLayout Layout { get; set; }

        public bool Overwrite { get; set; }
        public bool WelcomePage { get; set; }

        public List<WebPart> WebParts
        {
            get { return _webParts; }
            private set { _webParts = value; }
        }

        /// <summary>
        /// Defines the Security rules for the Page
        /// </summary>
        public ObjectSecurity Security
        {
            get { return this._security; }
            private set { this._security = value; }
        }

        #endregion

        #region Constructors
        public Page() { }

        public Page(string url, bool overwrite, WikiPageLayout layout, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null)
        {
            this.Url = url;
            this.Overwrite = overwrite;
            this.Layout = layout;
            this.WelcomePage = welcomePage;

            if (webParts != null)
            {
                this.WebParts.AddRange(webParts);
            }

            if (security != null)
            {
                this.Security = security;
            }
        }


        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                (this.Url != null ? this.Url.GetHashCode() : 0),
                this.Overwrite.GetHashCode(),
                this.Layout.GetHashCode()
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

        public bool Equals(Page other)
        {
            return (this.Url == other.Url &&
                this.Overwrite == other.Overwrite &&
                this.Layout == other.Layout);
        }

        #endregion
    }
}
