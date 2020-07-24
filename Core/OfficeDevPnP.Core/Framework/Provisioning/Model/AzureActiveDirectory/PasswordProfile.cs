using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.AzureActiveDirectory
{
    /// <summary>
    /// Defines the Password Profile for a User in AAD
    /// </summary>
    public partial class PasswordProfile : BaseModel, IEquatable<PasswordProfile>
    {
        #region Public Members

        /// <summary>
        /// Defines whether to force password change at next sign-in for the user
        /// </summary>
        public Boolean ForceChangePasswordNextSignIn { get; set; }

        /// <summary>
        /// Defines whether to force password change at next sign-in with MFA for the user
        /// </summary>
        public Boolean ForceChangePasswordNextSignInWithMfa { get; set; }

        /// <summary>
        /// The Password for the user
        /// </summary>
        public SecureString Password { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                ForceChangePasswordNextSignIn.GetHashCode(),
                ForceChangePasswordNextSignInWithMfa.GetHashCode(),
                Password.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with PasswordProfile class
        /// </summary>
        /// <param name="obj">Object that represents PasswordProfile</param>
        /// <returns>Checks whether object is PasswordProfile class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is PasswordProfile))
            {
                return (false);
            }
            return (Equals((PasswordProfile)obj));
        }

        /// <summary>
        /// Compares PasswordProfile object based on PackagePath and source
        /// </summary>
        /// <param name="other">PasswordProfile Class object</param>
        /// <returns>true if the PasswordProfile object is equal to the current object; otherwise, false.</returns>
        public bool Equals(PasswordProfile other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ForceChangePasswordNextSignIn == other.ForceChangePasswordNextSignIn &&
                this.ForceChangePasswordNextSignInWithMfa == other.ForceChangePasswordNextSignInWithMfa &&
                this.Password == other.Password
                );
        }

        #endregion
    }
}
