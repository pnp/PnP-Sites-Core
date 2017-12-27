using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Quicklaunch Navigation settings of a site
    /// </summary>
    public partial class Quicklaunch : BaseModel, IEquatable<Quicklaunch>
    {
        #region Public Members

        /// <summary>
        /// Defines whether to remove existing nodes before creating those described through this element
        /// </summary>
        public Boolean RemoveExistingNodes { get; set; }

        /// <summary>
        /// A collection of navigation nodes for the site
        /// </summary>
        public NavigationNodeCollection NavigationNodes { get; private set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for Quicklaunch class
        /// </summary>
        public Quicklaunch()
        {
            this.NavigationNodes = new NavigationNodeCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                this.NavigationNodes.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.RemoveExistingNodes.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Quicklaunch
        /// </summary>
        /// <param name="obj">Object that represents Quicklaunch</param>
        /// <returns>true if the current object is equal to the Quicklaunch</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Quicklaunch))
            {
                return (false);
            }
            return (Equals((Quicklaunch)obj));
        }

        /// <summary>
        /// Compares Quicklaunch object based on NavigationNodes and RemoveExistingNodes properties.
        /// </summary>
        /// <param name="other">Quicklaunch object</param>
        /// <returns>true if the Quicklaunch object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Quicklaunch other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.NavigationNodes.DeepEquals(other.NavigationNodes) &&
                this.RemoveExistingNodes == other.RemoveExistingNodes
                );
        }

        #endregion
    }
}
