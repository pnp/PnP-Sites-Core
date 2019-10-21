using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Base abstract class for the navigation kinds (global or current)
    /// </summary>
    public abstract partial class BaseNavigationKind : BaseModel, IEquatable<BaseNavigationKind>
    {
        #region Private Fields

        private StructuralNavigation _structuralNavigation;
        private ManagedNavigation _managedNavigation;

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the Structural Navigation settings of the site
        /// </summary>
        public StructuralNavigation StructuralNavigation
        {
            get { return (this._structuralNavigation); }
            private set
            {
                if (this._structuralNavigation != null)
                {
                    this._structuralNavigation.ParentTemplate = null;
                }
                this._structuralNavigation = value;
                if (this._structuralNavigation != null)
                {
                    this._structuralNavigation.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines the Managed Navigation settings of the site
        /// </summary>
        public ManagedNavigation ManagedNavigation
        {
            get { return (this._managedNavigation); }
            private set
            {
                if (this._managedNavigation != null)
                {
                    this._managedNavigation.ParentTemplate = null;
                }
                this._managedNavigation = value;
                if (this._managedNavigation != null)
                {
                    this._managedNavigation.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for BaseNavigationKind class
        /// </summary>
        public BaseNavigationKind()
        {

        }

        /// <summary>
        /// Constructor for BaseNavigationKind class
        /// </summary>
        /// <param name="structuralNavigation">Structural Navigation object</param>
        /// <param name="managedNavigation">Managed Navigation object</param>
        public BaseNavigationKind(StructuralNavigation structuralNavigation = null, ManagedNavigation managedNavigation = null)
        {
            this.StructuralNavigation = structuralNavigation;
            this.ManagedNavigation = managedNavigation;
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code.
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                this.StructuralNavigation.GetHashCode(),
                this.ManagedNavigation.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with BaseNavigationKind
        /// </summary>
        /// <param name="obj">Object that represents BaseNavigationKind</param>
        /// <returns>true if the current object is equal to the BaseNavigationKind</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is BaseNavigationKind))
            {
                return (false);
            }
            return (Equals((BaseNavigationKind)obj));
        }

        /// <summary>
        /// Compares BaseNavigationKind object based on StructuralNavigation and ManagedNavigation
        /// </summary>
        /// <param name="other">BaseNavigationKind object</param>
        /// <returns>true if the BaseNavigationKind object is equal to the current object; otherwise, false.</returns>
        public bool Equals(BaseNavigationKind other)
        {
            if (other == null)
            {
                return (false);
            }

            return (((this.StructuralNavigation != null && other.StructuralNavigation != null) ? this.StructuralNavigation.Equals(other.StructuralNavigation) : (this.StructuralNavigation == null && other.StructuralNavigation == null)) &&
                    ((this.ManagedNavigation != null && other.ManagedNavigation != null) ? this.ManagedNavigation == other.ManagedNavigation : (this.ManagedNavigation == null && other.ManagedNavigation == null))
                );
        }

        #endregion
    }
}
