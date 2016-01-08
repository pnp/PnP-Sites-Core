using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class File : BaseModel, IEquatable<File>
    {
        #region Private Members

        private WebPartCollection _webParts;
        private Dictionary<string, string> _properties = new Dictionary<string, string>();
        private ObjectSecurity _security;

        #endregion

        #region Properties
        public string Src { get; set; }

        public string Folder { get; set; }

        public bool Overwrite { get; set; }

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
        /// Defines the Security rules for the File
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
        public File()
        {
            this._webParts = new WebPartCollection(this.ParentTemplate);
        }

        public File(string src, string folder, bool overwrite, IEnumerable<WebPart> webParts, IDictionary<string, string> properties, ObjectSecurity security = null) :
            this()
        {
            this.Src = src;
            this.Overwrite = overwrite;
            this.Folder = folder;
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
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}",
                (this.Folder != null ? this.Folder.GetHashCode() : 0),
                this.Overwrite.GetHashCode(),
                (this.Src != null ? this.Src.GetHashCode() : 0),
                this.WebParts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.Security != null ? this.Security.GetHashCode() : 0)
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

        public bool Equals(File other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Folder == other.Folder &&
                this.Overwrite == other.Overwrite &&
                this.Src == other.Src &&
                this.WebParts.DeepEquals(other.WebParts) &&
                this.Properties.DeepEquals(other.Properties) &&
                (this.Security != null ? this.Security.Equals(other.Security) : true)
            );
        }

        #endregion
    }
}
