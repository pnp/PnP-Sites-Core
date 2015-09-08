using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class File : IEquatable<File>
    {
        #region Private Members

        private List<WebPart> _webParts = new List<WebPart>();
        private Dictionary<string, string> _properties = new Dictionary<string,string>();
        private ObjectSecurity _security = new ObjectSecurity();

        #endregion

        #region Properties
        public string Src { get; set; }

        public string Folder { get; set; }

        public bool Overwrite { get; set; }

        public List<WebPart> WebParts
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
            set { this._security = value; }
        }

        #endregion

        #region Constructors
        public File() { }

        public File(string src, string folder, bool overwrite, IEnumerable<WebPart> webParts, IDictionary<string,string> properties, ObjectSecurity security = null)
        {
            this.Src = src;
            this.Overwrite = overwrite;
            this.Folder = folder;
            if (webParts != null)
            {
                this.WebParts.AddRange(webParts);
            }
            if (properties != null)
            {
                foreach (var property in properties)
                {
                    this.Properties.Add(property.Key,property.Value);
                }
            }
            if (security != null)
            {
                this._security = security;
            }
        }


        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.Folder.GetHashCode(),
                this.Overwrite.GetHashCode(),
                this.Src.GetHashCode()
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
            return (this.Folder == other.Folder &&
                this.Overwrite == other.Overwrite &&
                this.Src == other.Src);
        }

        #endregion
    }
}
