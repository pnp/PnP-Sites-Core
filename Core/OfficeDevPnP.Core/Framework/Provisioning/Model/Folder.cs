using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a folder that will be provisioned into the target list/library
    /// </summary>
    public partial class Folder : BaseModel, IEquatable<Folder>
    {
        #region Private members

        private ObjectSecurity _objectSecurity;
        private FolderCollection _folders;
        private PropertyBagEntryCollection _propertyBags;

        #endregion

        #region Properties

        /// <summary>
        /// The Name of the Folder
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Defines the security rules for the current Folder
        /// </summary>
        public ObjectSecurity Security
        {
            get { return _objectSecurity; }
            private set
            {
                if (this._objectSecurity != null)
                {
                    this._objectSecurity.ParentTemplate = null;
                }
                this._objectSecurity = value;
                if (this._objectSecurity != null)
                {
                    this._objectSecurity.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines the child folders of the current Folder, if any
        /// </summary>
        public FolderCollection Folders
        {
            get { return _folders; }
            private set { _folders = value; }
        }

        public PropertyBagEntryCollection PropertyBagEntries
        {
            get { return this._propertyBags; }
            private set { this._propertyBags = value; }
        }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for the Folder class
        /// </summary>
        public Folder()
        {
            this.Security = new ObjectSecurity();
            this._folders = new FolderCollection(this.ParentTemplate);
            this._propertyBags = new PropertyBagEntryCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for the Folder class
        /// </summary>
        /// <param name="name">Name of the folder</param>
        /// <param name="folders">List of the folders</param>
        /// <param name="security">ObjectSecurity for the folder</param>
        public Folder(String name, List<Folder> folders = null, ObjectSecurity security = null) :
            this()
        {
            this.Name = name;
            this.Folders.AddRange(folders);
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
            return (String.Format("{0}|{1}|{2}|{3}|",
                (this.Name.GetHashCode()),
                (this.Folders.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))),
                (this.Security != null ? this.Security.GetHashCode() : 0),
                this.PropertyBagEntries.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Folder
        /// </summary>
        /// <param name="obj">Object that represents Folder</param>
        /// <returns>true if the current object is equal to the Folder</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Folder))
            {
                return (false);
            }
            return (Equals((Folder)obj));
        }

        /// <summary>
        /// Compares Folder object based on Name, Folders and Security properties.
        /// </summary>
        /// <param name="other">Folder object</param>
        /// <returns>true if the Folder object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Folder other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                    this.Folders.DeepEquals(other.Folders) &&
                    (this.Security != null ? this.Security.Equals(other.Security) : true) &&
                    this.PropertyBagEntries.DeepEquals(other.PropertyBagEntries)
               );
        }

        #endregion
    }
}
