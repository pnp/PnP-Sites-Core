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
    public partial class Folder : IEquatable<Folder>
    {
        #region Private members

        private ObjectSecurity _objectSecurity = new ObjectSecurity();
        private List<Folder> _folders = new List<Folder>();

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
            private set { _objectSecurity = value; }
        }

        /// <summary>
        /// Defines the child folders of the current Folder, if any
        /// </summary>
        public List<Folder> Folders
        {
            get { return _folders; }
            private set { _folders = value; }
        }

        #endregion

        #region Constructors

        public Folder() { }

        public Folder(String name, List<Folder> folders = null, ObjectSecurity security = null)
        {
            this.Name = name;

            if (folders != null)
            {
                this.Folders = folders;
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
                (this.Name.GetHashCode()),
                (this.Folders.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))),
                this.Security.GetHashCode()
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Folder))
            {
                return (false);
            }
            return (Equals((Folder)obj));
        }

        public bool Equals(Folder other)
        {
            return (this.Name == other.Name &&
                    this.Folders.DeepEquals(other.Folders) &&
                    this.Security.Equals(other.Security)
                );
        }

        #endregion
    }
}
