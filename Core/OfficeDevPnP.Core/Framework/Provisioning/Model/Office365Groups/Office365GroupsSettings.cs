using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Office365Groups
{
    /// <summary>
    /// Defines a Office365GroupsSettings object
    /// </summary>
    public partial class Office365GroupsSettings : BaseModel, IEquatable<Office365GroupsSettings>
    {
        #region Public members

        /// <summary>
        /// Properties of the file
        /// </summary>
        public Dictionary<string, string> Properties { get; private set; } = new Dictionary<string, string>();

        #endregion

        #region Constructors

        public Office365GroupsSettings() : base()
        {
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}",
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Office365GroupsSettings class
        /// </summary>
        /// <param name="obj">Object that represents Office365GroupsSettings</param>
        /// <returns>Checks whether object is Office365GroupsSettings class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Office365GroupsSettings))
            {
                return (false);
            }
            return (Equals((Office365GroupsSettings)obj));
        }

        /// <summary>
        /// Compares Office365GroupsSettings object based on DriveRoots
        /// </summary>
        /// <param name="other">Office365GroupsSettings Class object</param>
        /// <returns>true if the Office365GroupsSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Office365GroupsSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.Properties.DeepEquals(other.Properties)
                );
        }

        #endregion
    }
}
