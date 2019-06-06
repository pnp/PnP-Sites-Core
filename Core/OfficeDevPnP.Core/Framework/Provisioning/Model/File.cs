using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a File element, to describe a file that will be provisioned into the target Site
    /// </summary>
    public partial class File : BaseModel, IEquatable<File>
    {
        #region Private Members

        private WebPartCollection _webParts;
        private Dictionary<string, string> _properties = new Dictionary<string, string>();
        private ObjectSecurity _security;

        #endregion

        #region Properties

        /// <summary>
        /// The Src of the File
        /// </summary>
        public string Src { get; set; }

        /// <summary>
        /// The TargetFolder of the File
        /// </summary>
        public string Folder { get; set; }

        /// <summary>
        /// The Overwrite flag for the File
        /// </summary>
        public bool Overwrite { get; set; }

        /// <summary>
        /// The Level status for the File
        /// </summary>
        public FileLevel Level { get; set; }

        /// <summary>
        /// Webparts in the file
        /// </summary>
        public WebPartCollection WebParts
        {
            get { return _webParts; }
            private set { _webParts = value; }
        }

        /// <summary>
        /// Properties of the file
        /// </summary>
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

        /// <summary>
        /// The Target file name for the File, optional attribute. If missing, the original file name will be used.
        /// </summary>
        public String TargetFileName { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for the File class
        /// </summary>
        public File()
        {
            this._webParts = new WebPartCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for the File class
        /// </summary>
        /// <param name="src">Source name of the file</param>
        /// <param name="folder">Targer Folder of the file</param>
        /// <param name="overwrite">Overwrite flag of the file</param>
        /// <param name="webParts">Webparts in the file</param>
        /// <param name="properties">Properties of the file</param>
        /// <param name="security">Security Rules of the file</param>
        /// <param name="level">Level status for the file</param>
        public File(string src, string folder, bool overwrite, IEnumerable<WebPart> webParts, IDictionary<string, string> properties, ObjectSecurity security = null, FileLevel level = FileLevel.Draft) :
            this()
        {
            this.Src = src;
            this.Overwrite = overwrite;
            this.Level = level;
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
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|",
                (this.Folder != null ? this.Folder.GetHashCode() : 0),
                this.Overwrite.GetHashCode(),
                (this.Src != null ? this.Src.GetHashCode() : 0),
                this.WebParts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.Security != null ? this.Security.GetHashCode() : 0),
                this.TargetFileName?.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with File
        /// </summary>
        /// <param name="obj">Object that represents File</param>
        /// <returns>true if the current object is equal to the File</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is File))
            {
                return (false);
            }
            return (Equals((File)obj));
        }

        /// <summary>
        /// Compares File object based on Folder, Overwrite, Src, WebParts, Properties, Security, and TargetFileName.
        /// </summary>
        /// <param name="other">File object</param>
        /// <returns>true if the File object is equal to the current object; otherwise, false.</returns>
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
                (this.Security != null ? this.Security.Equals(other.Security) : true) &&
                this.TargetFileName == other.TargetFileName
            );
        }

        #endregion
    }
}
