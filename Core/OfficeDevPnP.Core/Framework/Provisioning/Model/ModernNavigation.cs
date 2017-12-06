using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// The Modern Navigation settings for the Provisioning Template
    /// </summary>
    public partial class ModernNavigation : BaseNavigationKind, IEquatable<ModernNavigation>
    {
        #region Public Members


        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for ModernNavigation class
        /// </summary>
        public ModernNavigation()
        {

        }

        /// <summary>
        /// Constructor for ModernNavigation class
        /// </summary>
        /// <param name="quicklaunch"></param>
        public ModernNavigation(Quicklaunch quicklaunch):
            base(null, null, quicklaunch)
        {
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
                base.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ModernNavigation
        /// </summary>
        /// <param name="obj">Object that represents ModernNavigation</param>
        /// <returns>true if the Modern object is equal to the ModernNavigation</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ModernNavigation))
            {
                return (false);
            }
            return (Equals((ModernNavigation)obj));
        }

        /// <summary>
        /// Compares ModernNavigation object based on BaseNavigationKind and NavigationType.
        /// </summary>
        /// <param name="other">ModernNavigation object</param>
        /// <returns>true if the ModernNavigation object is equal to the Modern object; otherwise, false.</returns>
        public bool Equals(ModernNavigation other)
        {
            if (other == null)
            {
                return (false);
            }

            return ((BaseNavigationKind)this == (BaseNavigationKind)other);
        }

        #endregion
    }

    ///// <summary>
    ///// Defines the type of Modern Navigation
    ///// </summary>
    //public enum ModernNavigationType
    //{
    //    /// <summary>
    //    /// The site inherits the Global Navigation settings from its parent
    //    /// </summary>
    //    Inherit,
    //    /// <summary>
    //    /// The site uses Structural Global Navigation
    //    /// </summary>
    //    Structural,
    //    /// <summary>
    //    /// The site uses Structural Local Modern Navigation
    //    /// </summary>
    //    StructuralLocal,
    //    /// <summary>
    //    /// The site uses Managed Global Navigation
    //    /// </summary>
    //    Managed,
    //}
}
