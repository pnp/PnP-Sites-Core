using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Structural Navigation settings of a site
    /// </summary>
    public partial class StructuralNavigation : BaseModel, IEquatable<StructuralNavigation>
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
        /// Constructor for StructuralNavigation class
        /// </summary>
        public StructuralNavigation()
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
        /// Compares object with StructuralNavigation
        /// </summary>
        /// <param name="obj">Object that represents StructuralNavigation</param>
        /// <returns>true if the current object is equal to the StructuralNavigation</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is StructuralNavigation))
            {
                return (false);
            }
            return (Equals((StructuralNavigation)obj));
        }

        /// <summary>
        /// Compares StructuralNavigation object based on NavigationNodes and RemoveExistingNodes properties.
        /// </summary>
        /// <param name="other">StructuralNavigation object</param>
        /// <returns>true if the StructuralNavigation object is equal to the current object; otherwise, false.</returns>
        public bool Equals(StructuralNavigation other)
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
