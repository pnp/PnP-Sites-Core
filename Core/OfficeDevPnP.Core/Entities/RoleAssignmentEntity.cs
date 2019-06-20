using System;
using System.Collections.Generic;
using System.Text;
namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Defines the role assignments for a user or group.
    /// </summary>
    public class RoleAssignmentEntity
    {
        /// <summary>
        /// Path
        /// </summary>
        public string Path
        {
            get;
            set;
        }

        /// <summary>
        /// User entity
        /// </summary>
        public UserEntity User
        {
            get;
            set;
        }

        /// <summary>
        /// User Title
        /// </summary>
        public string UserTitle
        {
            get
            {
                return User.Title;
            }
            set
            {
            }
        }

        /// <summary>
        /// User login name
        /// </summary>
        public string UserLoginName
        {
            get
            {
                return User.LoginName;
            }
            set
            {
            }
        }

        /// <summary>
        /// User email
        /// </summary>
        public string UserEmail
        {
            get
            {
                return string.IsNullOrWhiteSpace(User.Email) ? null : User.Email;
            }
            set
            {
            }
        }

        /// <summary>
        /// User role
        /// </summary>
        public string Role
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the collection of role definition bindings for the role assignment
        /// </summary>
        public ICollection<String> RoleDefinitionBindings
        {
            get;
            set;
        }

        /// <summary>
        /// user permissions
        /// </summary>
        public string Permissions
        {
            get
            {
                return string.Join(";", RoleDefinitionBindings);
            }
            set
            {
            }
        }

        /// <summary>
        /// Tag for the user
        /// </summary>
        public string Tag
        {
            get;
            set;
        }

        /// <summary>
        /// DateTime value of RoleAssignment created
        /// </summary>
        public DateTime CreatedDate
        {
            get;
            set;
        }

        /// <summary>
        /// Returns a string that represents the current object
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append(Path);
            buffer.Append("\t");
            buffer.Append(User.Title);
            buffer.Append("\t");
            buffer.Append(string.IsNullOrWhiteSpace(User.Email) ? "(n/a)" : User.Email);
            buffer.Append("\t");
            buffer.Append(User.LoginName);
            buffer.Append("\t");
            buffer.Append(Role);
            buffer.Append("\t");
            buffer.Append(string.Join(";", RoleDefinitionBindings));
            return buffer.ToString();
        }
    }
}
