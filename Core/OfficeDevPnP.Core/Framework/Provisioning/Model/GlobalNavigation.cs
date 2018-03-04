using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// The Global Navigation settings for the Provisioning Template
    /// </summary>
    public partial class GlobalNavigation : BaseNavigationKind, IEquatable<GlobalNavigation>
    {
        #region Public Members

        /// <summary>
        /// Defines the type of Global Navigation
        /// </summary>
        public GlobalNavigationType NavigationType { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for GlobalNavigation class
        /// </summary>
        public GlobalNavigation()
        {

        }

        /// <summary>
        /// Constructor for GlobalNavigation class
        /// </summary>
        /// <param name="navigationType">Global Navigation Type</param>
        /// <param name="structuralNavigation">StructuralNavigation object</param>
        /// <param name="managedNavigation">ManagedNavigation object</param>
        public GlobalNavigation(GlobalNavigationType navigationType, StructuralNavigation structuralNavigation = null, ManagedNavigation managedNavigation = null):
            base(structuralNavigation, managedNavigation)
        {
            this.NavigationType = navigationType;
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
                base.GetHashCode(),
                this.NavigationType.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with GlobalNavigation
        /// </summary>
        /// <param name="obj">Object that represents GlobalNavigation</param>
        /// <returns>true if the current object is equal to the GlobalNavigation</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is GlobalNavigation))
            {
                return (false);
            }
            return (Equals((GlobalNavigation)obj));
        }

        /// <summary>
        /// Compares GlobalNavigation object based on BaseNavigationKind object and NavigationType property.
        /// </summary>
        /// <param name="other">GlobalNavigation object</param>
        /// <returns>true if the GlobalNavigation object is equal to the current object; otherwise, false.</returns>
        public bool Equals(GlobalNavigation other)
        {
            if (other == null)
            {
                return (false);
            }

            return (((BaseNavigationKind)this).Equals((BaseNavigationKind)other) &&
                this.NavigationType == other.NavigationType
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the type of Global Navigation
    /// </summary>
    public enum GlobalNavigationType
    {
        /// <summary>
        /// The site inherits the Global Navigation settings from its parent
        /// </summary>
        Inherit,
        /// <summary>
        /// The site uses Structural Global Navigation
        /// </summary>
        Structural,
        /// <summary>
        /// The site uses Managed Global Navigation
        /// </summary>
        Managed,
    }
}
