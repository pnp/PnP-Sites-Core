using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines an available Page Layout for the current Publishing site
    /// </summary>
    public partial class PageLayout : BaseModel, IEquatable<PageLayout>
    {
        #region Public Members

        /// <summary>
        /// Defines the path of the Page Layout for the current Publishing site
        /// </summary>
        public String Path { get; set; }

        /// <summary>
        /// Defines whether the Page Layout is the default for the current Publishing site
        /// </summary>
        public Boolean IsDefault { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                (this.Path != null ? this.Path.GetHashCode() : 0),
                this.IsDefault.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with PageLayout
        /// </summary>
        /// <param name="obj">Object that represents PageLayout</param>
        /// <returns>true if the current object is equal to the PageLayout</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is PageLayout))
            {
                return (false);
            }
            return (Equals((PageLayout)obj));
        }

        /// <summary>
        /// Compares PageLayout object based on Path and IsDefault properties.
        /// </summary>
        /// <param name="other">PageLayout object</param>
        /// <returns>true if the PageLayout object is equal to the current object; otherwise, false.</returns>
        public bool Equals(PageLayout other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.Path == other.Path &&
                this.IsDefault == other.IsDefault
                );
        }

        #endregion
    }
}
