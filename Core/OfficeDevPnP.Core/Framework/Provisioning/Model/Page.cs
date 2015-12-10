using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class Page : IEquatable<Page>
    {
        #region Private Members

        private List<WebPart> _webParts = new List<WebPart>();
        private ObjectSecurity _security = null;
        private Dictionary<String, String> _fields = new Dictionary<String, String>();

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

        /// <summary>
        /// The Fields to setup for the Page
        /// </summary>
        public Dictionary<String, String> Fields
        {
            get { return this._fields; }
            private set { this._fields = value; }
        }

        #endregion

        #region Constructors
        public Page() { }

        public Page(string url, bool overwrite, WikiPageLayout layout, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null):
            this(url, overwrite, layout, webParts, welcomePage, security, null)
        {
        }

        public Page(string url, bool overwrite, WikiPageLayout layout, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null, Dictionary<String, String> fields = null)
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

            if (fields != null)
            {
                this.Fields = fields;
            }
        }


        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|",
                (this.Url != null ? this.Url.GetHashCode() : 0),
                this.Overwrite.GetHashCode(),
                this.Layout.GetHashCode(),
                this.WelcomePage.GetHashCode(),
                this.WebParts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Security.GetHashCode(),
                this.Fields.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Page))
            {
                return (false);
            }
            return (Equals((Page)obj));
        }

        public bool Equals(Page other)
        {
            return (this.Url == other.Url &&
                this.Overwrite == other.Overwrite &&
                this.Layout == other.Layout &&
                this.WelcomePage == other.WelcomePage &&
                this.WebParts.DeepEquals(other.WebParts) &&
                this.Security.Equals(other.Security) &&
                this.Fields.DeepEquals(other.Fields)
                );
        }

        #endregion
    }
}
