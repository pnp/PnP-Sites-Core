using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object used in the Provisioning template that defines a Section of Settings for the current Site Collection
    /// </summary>
    public partial class SiteSettings : BaseModel, IEquatable<SiteSettings>
    {
        #region Properties

        /// <summary>
        /// Defines whether a designer can be used on this site collection
        /// </summary>
        public Boolean AllowDesigner { get; set; }

        /// <summary>
        /// Defines whether creation of declarative workflows is allowed in the site collection
        /// </summary>
        public Boolean AllowCreateDeclarativeWorkflow { get; set; }

        /// <summary>
        /// Defines whether saving of declarative workflows is allowed in the site collection
        /// </summary>
        public Boolean AllowSaveDeclarativeWorkflowAsTemplate { get; set; }

        /// <summary>
        /// Defines whether publishing of declarative workflows is allowed in the site collection
        /// </summary>
        public Boolean AllowSavePublishDeclarativeWorkflow { get; set; }

        /// <summary>
        /// Defines whether social bar is disabled on Site Pages in this site collection
        /// </summary>
        public Boolean SocialBarOnSitePagesDisabled { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Default Constructor
        /// </summary>
        public SiteSettings() { }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets hash code
        /// </summary>
        /// <returns>Returns hash code in integer</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|",
                (this.AllowDesigner.GetHashCode()),
                (this.AllowCreateDeclarativeWorkflow.GetHashCode()),
                (this.AllowSaveDeclarativeWorkflowAsTemplate.GetHashCode()),
                (this.AllowSavePublishDeclarativeWorkflow.GetHashCode()),
                (this.SocialBarOnSitePagesDisabled.GetHashCode())
            ).GetHashCode());
        }

        /// <summary>
        /// Compares web settings with other web settings
        /// </summary>
        /// <param name="obj">SiteSettings object</param>
        /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteSettings))
            {
                return (false);
            }
            return (Equals((SiteSettings)obj));
        }

        /// <summary>
        /// Compares SiteSettings with other web settings
        /// </summary>
        /// <param name="other">SiteSettings object</param>
        /// <returns>true if the SiteSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AllowDesigner == other.AllowDesigner &&
                    this.AllowCreateDeclarativeWorkflow == other.AllowCreateDeclarativeWorkflow &&
                    this.AllowSaveDeclarativeWorkflowAsTemplate == other.AllowSaveDeclarativeWorkflowAsTemplate &&
                    this.AllowSavePublishDeclarativeWorkflow == other.AllowSavePublishDeclarativeWorkflow &&
                    this.SocialBarOnSitePagesDisabled == other.SocialBarOnSitePagesDisabled
                );
        }

        #endregion
    }
}
