using System;
using System.Collections.Generic;
using System.Text;
namespace OfficeDevPnP.Core.Entities
{
    public class RoleAssignmentEntity
    {
        public string Path
        {
            get;
            set;
        }

        public UserEntity User
        {
            get;
            set;
        }

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

        public string Role
        {
            get;
            set;
        }

        public ICollection<String> RoleDefinitionBindings
        {
            get;
            set;
        }

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

        public string Tag
        {
            get;
            set;
        }

        public DateTime CreatedDate
        {
            get;
            set;
        }

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
