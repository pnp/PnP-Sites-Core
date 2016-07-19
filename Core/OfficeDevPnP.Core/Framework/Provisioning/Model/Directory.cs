using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Directory element, to describe a folder in the current 
    /// repository that will be used to upload files into the target Site
    /// </summary>
    public partial class Directory : BaseModel, IEquatable<Directory>
    {
        #region Private Members

        private ObjectSecurity _security;

        #endregion

        #region Properties

        /// <summary>
        /// The Src of the Directory
        /// </summary>
        public string Src { get; set; }

        /// <summary>
        /// The TargetFolder of the Directory
        /// </summary>
        public string Folder { get; set; }

        /// <summary>
        /// The Overwrite flag for the files in the Directory
        /// </summary>
        public bool Overwrite { get; set; }

        /// <summary>
        /// The Level status for the files in the Directory
        /// </summary>
        public FileLevel Level { get; set; }

        /// <summary>
        /// Defines whether to recursively browse through all the child folders of the Directory
        /// </summary>
        public bool Recursive { get; set; }

        /// <summary>
        /// The file Extensions to include while uploading the Directory
        /// </summary>
        public String IncludedExtensions { get; set; }

        /// <summary>
        /// The file Extensions to exclude while uploading the Directory
        /// </summary>
        public String ExcludedExtensions { get; set; }

        /// <summary>
        /// The file path of JSON mapping file with metadata for files to upload in the Directory
        /// </summary>
        public String MetadataMappingFile { get; set; }

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
        public Directory()
        {
        }

        public Directory(string src, string folder, bool overwrite, 
            FileLevel level = FileLevel.Draft, bool recursive = false, 
            string includeExtensions = null, string excludeExtensions = null, 
            string metadataMappingFile = null, ObjectSecurity security = null) :
            this()
        {
            this.Src = src;
            this.Folder = folder;
            this.Overwrite = overwrite;
            this.Level = level;
            this.Recursive = recursive;
            this.IncludedExtensions = includeExtensions;
            this.ExcludedExtensions = excludeExtensions;
            this.MetadataMappingFile = metadataMappingFile;
            if (security != null)
            {
                this.Security = security;
            }
        }


        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}",
                (this.ExcludedExtensions != null ? this.ExcludedExtensions.GetHashCode() : 0),
                (this.Folder != null ? this.Folder.GetHashCode() : 0),
                (this.IncludedExtensions != null ? this.IncludedExtensions.GetHashCode() : 0),
                this.Level.GetHashCode(),
                (this.MetadataMappingFile != null ? this.MetadataMappingFile.GetHashCode() : 0),
                this.Overwrite.GetHashCode(),
                this.Recursive.GetHashCode(),
                (this.Src != null ? this.Src.GetHashCode() : 0),
                (this.Security != null ? this.Security.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Directory))
            {
                return (false);
            }
            return (Equals((Directory)obj));
        }

        public bool Equals(Directory other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ExcludedExtensions == other.ExcludedExtensions &&
                this.Folder == other.Folder &&
                this.IncludedExtensions == other.IncludedExtensions &&
                this.Level == other.Level &&
                this.MetadataMappingFile == other.MetadataMappingFile &&
                this.Overwrite == other.Overwrite &&
                this.Recursive == other.Recursive &&
                this.Src == other.Src &&
                (this.Security != null ? this.Security.Equals(other.Security) : true)
            );
        }

        #endregion
    }
}
