using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Navigation Node for the Structural Navigation of a site
    /// </summary>
    public partial class NavigationNode : BaseModel, IEquatable<NavigationNode>
    {
        #region Public Members

        /// <summary>
        /// A collection of navigation nodes children of the current NavigatioNode
        /// </summary>
        public NavigationNodeCollection NavigationNodes { get; private set; }

        /// <summary>
        /// Defines the Title of a Navigation Node
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Defines the Url of a Navigation Node
        /// </summary>
        public String Url { get; set; }

        /// <summary>
        /// Defines whether the Navigation Node for the Structural Navigation targets an External resource
        /// </summary>
        public Boolean IsExternal { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for NavigationNode class
        /// </summary>
        public NavigationNode()
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
            return (String.Format("{0}|{1}|{2}|{3}",
                this.IsExternal.GetHashCode(),
                this.NavigationNodes.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                (this.Title != null ? this.Title.GetHashCode() : 0),
                (this.Url != null ? this.Url.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with NavigationNode
        /// </summary>
        /// <param name="obj">Object that represents NavigationNode</param>
        /// <returns>true if the current object is equal to the NavigationNode</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is NavigationNode))
            {
                return (false);
            }
            return (Equals((NavigationNode)obj));
        }

        /// <summary>
        /// Compares NavigationNode object based on isExternal, NavigationNodes, Title and Url properties.
        /// </summary>
        /// <param name="other">NavigationNode object</param>
        /// <returns>true if the NavigationNode object is equal to the current object; otherwise, false.</returns>
        public bool Equals(NavigationNode other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.IsExternal == other.IsExternal &&
                this.NavigationNodes.DeepEquals(other.NavigationNodes) &&
                this.Title == other.Title &&
                this.Url == other.Url
                );
        }

        #endregion
    }
}
