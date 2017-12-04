using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines an ALM settings to provision
    /// </summary>
    public partial class ApplicationLifecycleManagement : BaseModel, IEquatable<ApplicationLifecycleManagement>
    {
        #region Public Members

        /// <summary>
        /// Defines the AppCatalog settings for the current Site Collection
        /// </summary>
        public AppCatalog AppCatalog { get; set; }

        /// <summary>
        /// Defines the Apps for the current Site Collection
        /// </summary>
        public AppCollection Apps { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                AppCatalog?.GetHashCode() ?? 0,
                Apps?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ApplicationLifecycleManagement class
        /// </summary>
        /// <param name="obj">Object that represents ApplicationLifecycleManagement</param>
        /// <returns>Checks whether object is ApplicationLifecycleManagement class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ApplicationLifecycleManagement))
            {
                return (false);
            }
            return (Equals((ApplicationLifecycleManagement)obj));
        }

        /// <summary>
        /// Compares ApplicationLifecycleManagement object based on AppCatalog, and Apps
        /// </summary>
        /// <param name="other">ApplicationLifecycleManagement Class object</param>
        /// <returns>true if the ApplicationLifecycleManagement object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ApplicationLifecycleManagement other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AppCatalog == other.AppCatalog &&
                this.Apps.DeepEquals(other.Apps)
                );
        }

        #endregion
    }
}
