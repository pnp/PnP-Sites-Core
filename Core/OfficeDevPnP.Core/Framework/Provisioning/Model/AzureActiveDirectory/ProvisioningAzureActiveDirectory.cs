using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.AzureActiveDirectory
{
    /// <summary>
    /// Defines a complex type declaring settings for provisioning Azure Active Directory objects
    /// </summary>
    public partial class ProvisioningAzureActiveDirectory : BaseModel, IEquatable<ProvisioningAzureActiveDirectory>
    {
        #region Public members
        /// <summary>
        /// Defines a collection of users to create in AAD
        /// </summary>
        public UserCollection Users { get; private set; }

        #endregion

        #region Constructors

        public ProvisioningAzureActiveDirectory()
        {
            this.Users = new UserCollection(this.ParentTemplate);
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
                this.Users.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ProvisioningAzureActiveDirectory
        /// </summary>
        /// <param name="obj">Object that represents ProvisioningAzureActiveDirectory</param>
        /// <returns>true if the current object is equal to the ProvisioningAzureActiveDirectory</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningAzureActiveDirectory))
            {
                return (false);
            }
            return (Equals((ProvisioningAzureActiveDirectory)obj));
        }

        /// <summary>
        /// Compares ProvisioningAzureActiveDirectory object based on Users
        /// </summary>
        /// <param name="other">ProvisioningAzureActiveDirectory object</param>
        /// <returns>true if the ProvisioningAzureActiveDirectory object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ProvisioningAzureActiveDirectory other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Users.DeepEquals(other.Users));
        }

        #endregion
    }
}
