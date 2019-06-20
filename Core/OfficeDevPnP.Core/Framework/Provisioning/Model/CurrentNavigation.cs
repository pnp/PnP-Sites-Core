using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// The Current Navigation settings for the Provisioning Template
    /// </summary>
    public partial class CurrentNavigation : BaseNavigationKind, IEquatable<CurrentNavigation>
    {
        #region Public Members

        /// <summary>
        /// Defines the type of Current Navigation
        /// </summary>
        public CurrentNavigationType NavigationType { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for CurrentNavigation class
        /// </summary>
        public CurrentNavigation()
        {

        }

        /// <summary>
        /// Constructor for CurrentNavigation class
        /// </summary>
        /// <param name="navigationType">CurrentNavigationType object</param>
        /// <param name="structuralNavigation">StructuralNavigation object</param>
        /// <param name="managedNavigation">ManagedNavigation object</param>
        public CurrentNavigation(CurrentNavigationType navigationType, StructuralNavigation structuralNavigation = null, ManagedNavigation managedNavigation = null):
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
        /// Compares object with CurrentNavigation
        /// </summary>
        /// <param name="obj">Object that represents CurrentNavigation</param>
        /// <returns>true if the current object is equal to the CurrentNavigation</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is CurrentNavigation))
            {
                return (false);
            }
            return (Equals((CurrentNavigation)obj));
        }

        /// <summary>
        /// Compares CurrentNavigation object based on BaseNavigationKind and NavigationType.
        /// </summary>
        /// <param name="other">CurrentNavigation object</param>
        /// <returns>true if the CurrentNavigation object is equal to the current object; otherwise, false.</returns>
        public bool Equals(CurrentNavigation other)
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
    /// Defines the type of Current Navigation
    /// </summary>
    public enum CurrentNavigationType
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
        /// The site uses Structural Local Current Navigation
        /// </summary>
        StructuralLocal,
        /// <summary>
        /// The site uses Managed Global Navigation
        /// </summary>
        Managed,
    }
}
