using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// The Navigation configurations of the Provisioning Template
    /// </summary>
    public partial class Navigation : BaseModel, IEquatable<Navigation>
    {
        #region Private Fields

        private GlobalNavigation _globalNavigation;
        private CurrentNavigation _currentNavigation;
        private ModernNavigation _modernNavigation;

        #endregion

        #region Public Members

        /// <summary>
        /// The Modern Navigation settings for the Provisioning Template
        /// </summary>
        public ModernNavigation ModernNavigation
        {
            get { return (this._modernNavigation);  }
            private set
            {
                if (this._modernNavigation != null)
                {
                    this._modernNavigation.ParentTemplate = null;
                }
                this._modernNavigation = value;
                if (this._modernNavigation != null)
                {
                    this._modernNavigation.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// The Global Navigation settings for the Provisioning Template
        /// </summary>
        public GlobalNavigation GlobalNavigation
        {
            get { return (this._globalNavigation); }
            private set
            {
                if (this._globalNavigation != null)
                {
                    this._globalNavigation.ParentTemplate = null;
                }
                this._globalNavigation = value;
                if (this._globalNavigation != null)
                {
                    this._globalNavigation.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// The Current Navigation settings for the Provisioning Template
        /// </summary>
        public CurrentNavigation CurrentNavigation
        {
            get { return (this._currentNavigation); }
            private set
            {
                if (this._currentNavigation != null)
                {
                    this._currentNavigation.ParentTemplate = null;
                }
                this._currentNavigation = value;
                if (this._currentNavigation != null)
                {
                    this._currentNavigation.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Declares whether the tree view has to be enabled at the site level or not, optional attribute.
        /// </summary>
        public Boolean EnableTreeView { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for Navigation class
        /// </summary>
        public Navigation()
        {

        }

        /// <summary>
        /// Constructor for Navigation class
        /// </summary>
        /// <param name="globalNavigation">GlobalNavigation object</param>
        /// <param name="currentNavigation">CurrentNavigation object</param>
        /// <param name="modernNavigation"></param>
        public Navigation(GlobalNavigation globalNavigation = null, CurrentNavigation currentNavigation = null, ModernNavigation modernNavigation = null)
        {
            this.GlobalNavigation = globalNavigation;
            this.CurrentNavigation = currentNavigation;
            this.ModernNavigation = modernNavigation;
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                (this.GlobalNavigation != null ? this.GlobalNavigation.GetHashCode() : 0),
                (this.CurrentNavigation != null ? this.CurrentNavigation.GetHashCode() : 0),
                this.EnableTreeView.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Navigation
        /// </summary>
        /// <param name="obj">Object that represents Navigation</param>
        /// <returns>true if the current object is equal to the Navigation</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Navigation))
            {
                return (false);
            }
            return (Equals((Navigation)obj));
        }

        /// <summary>
        /// Compares Navigation object based on GlobalNavigation and CurrentNavigation properties.
        /// </summary>
        /// <param name="other">Navigation object</param>
        /// <returns>true if the Navigation object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Navigation other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.GlobalNavigation == other.GlobalNavigation &&
                this.CurrentNavigation == other.CurrentNavigation &&
                this.EnableTreeView == other.EnableTreeView
                );
        }

        #endregion
    }
}
