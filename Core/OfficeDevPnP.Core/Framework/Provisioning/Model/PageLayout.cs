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

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                (this.Path != null ? this.Path.GetHashCode() : 0),
                this.IsDefault.GetHashCode()
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is PageLayout))
            {
                return (false);
            }
            return (Equals((PageLayout)obj));
        }

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
