using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a single tenant-wide Web API Permission
    /// </summary>
    public partial class WebApiPermission : BaseModel, IEquatable<WebApiPermission>
    {

        #region Public Members
        /// <summary>
        /// Gets or sets the Scope of the Web API permission
        /// </summary>
        public string Scope { get; set; }
        /// <summary>
        /// Gets or sets the target Resource of the Web API permission
        /// </summary>
        public string Resource { get; set; }

        #endregion

        #region Constructors
        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                this.Scope.GetHashCode(),
                (this.Resource != null ? this.Resource.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with WebApiPermission
        /// </summary>
        /// <param name="obj">Object that represents WebApiPermission</param>
        /// <returns>true if the current object is equal to the WebApiPermission</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is WebApiPermission))
            {
                return (false);
            }
            return (Equals((WebApiPermission)obj));
        }

        /// <summary>
        /// Compares WebApiPermission object based on Scope, Resource and Action. 
        /// </summary>
        /// <param name="other">WebApiPermission object</param>
        /// <returns>true if the WebApiPermission object is equal to the current object; otherwise, false.</returns>
        public bool Equals(WebApiPermission other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Scope == other.Scope &&
                this.Resource == other.Resource);
        }

        #endregion
    }
}
