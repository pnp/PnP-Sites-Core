using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class PublishingPage : BaseModel, IEquatable<PublishingPage>
    {
        #region Private Members

        private WebPartCollection _webParts;
        private Dictionary<string, string> _properties = new Dictionary<string, string>();
        private ObjectSecurity _security;

        #endregion

        #region Properties
        public string Name { get; set; }

        public string Layout { get; set; }

        public string Folder { get; set; }

        public bool Overwrite { get; set; }

        public bool Publish { get; set; }

        public WebPartCollection WebParts
        {
            get { return _webParts; }
            private set { _webParts = value; }
        }

        public Dictionary<string, string> Properties
        {
            get { return _properties; }
            private set { _properties = value; }
        }

        /// <summary>
        /// Defines the Security rules for the Publishing Page
        /// </summary>
        public ObjectSecurity Security
        {
            get { return this._security; }
            set
            {
                if (this._security != null)
                {
                    this._security.ParentTemplate = null;
                }
                this._security = value;
                if (this._security != null)
                {
                    this._security.ParentTemplate = this.ParentTemplate;
                }
            }
        }


        #endregion

        #region Constructors
        public PublishingPage()
        {
            this._webParts = new WebPartCollection(this.ParentTemplate);
        }

        public PublishingPage(string name, string layout, string folder, bool publish, bool overwrite, IEnumerable<WebPart> webParts, IDictionary<string, string> properties, ObjectSecurity security = null) :
            this()
        {
            this.Name = name;
            this.Layout = layout;
            this.Folder = folder;
            this.Publish = publish;
            this.Overwrite = overwrite;
            this.WebParts.AddRange(webParts);
            if (properties != null)
            {
                foreach (var property in properties)
                {
                    this.Properties.Add(property.Key, property.Value);
                }
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
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                (this.Folder != null ? this.Folder.GetHashCode() : 0),
                (this.Layout != null ? this.Layout.GetHashCode() : 0),
                this.Overwrite.GetHashCode(),
                this.Publish.GetHashCode(),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                this.WebParts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.Security != null ? this.Security.GetHashCode() : 0)
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

            return (this.Folder == other.Folder &&
                this.Layout == other.Layout &&
                this.Overwrite == other.Overwrite &&
                this.Publish == other.Publish &&
                this.Name == other.Name &&
                this.WebParts.DeepEquals(other.WebParts) &&
                this.Properties.DeepEquals(other.Properties) &&
                (this.Security != null ? this.Security.Equals(other.Security) : true)
            );
        }

        #endregion
    }
}
