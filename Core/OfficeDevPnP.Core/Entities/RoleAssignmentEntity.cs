using System;
using System.Collections.Generic;
using System.Text;
namespace OfficeDevPnP.Core.Entities
{
    //[Table("RoleAssignments")]
    public class RoleAssignmentEntity
    {
        //[Key]
        //public int Id
        //{
        //    get;
        //    set;
        //}

        public string Path
        {
            get;
            set;
        }

        //[NotMapped]
        public UserEntity User
        {
            get;
            set;
        }

        //[StringLength(256)]
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

        //[StringLength(256)]
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

        //[StringLength(256)]
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

        //[StringLength(256)]
        public string Role
        {
            get;
            set;
        }

        //[NotMapped]
        public ICollection<String> RoleDefinitionBindings
        {
            get;
            set;
        }

        //[StringLength(256)]
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

        //[StringLength(256)]
        public string Tag
        {
            get;
            set;
        }

        //[Column(TypeName = "datetime2")]
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
