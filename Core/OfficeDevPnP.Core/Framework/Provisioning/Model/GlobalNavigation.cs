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

        public GlobalNavigation()
        {

        }

        public GlobalNavigation(GlobalNavigationType navigationType, StructuralNavigation structuralNavigation = null, ManagedNavigation managedNavigation = null):
            base(structuralNavigation, managedNavigation)
        {
            this.NavigationType = navigationType;
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                base.GetHashCode(),
                this.NavigationType.GetHashCode()
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is GlobalNavigation))
            {
                return (false);
            }
            return (Equals((GlobalNavigation)obj));
        }

        public bool Equals(GlobalNavigation other)
        {
            if (other == null)
            {
                return (false);
            }

            return ((BaseNavigationKind)this == (BaseNavigationKind)other &&
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
