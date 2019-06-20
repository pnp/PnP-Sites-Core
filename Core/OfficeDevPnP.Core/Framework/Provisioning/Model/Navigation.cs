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
        private StructuralNavigation _searchNavigation;

        #endregion

        #region Public Members

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
        /// Defines the Search Navigation settings of the site
        /// </summary>
        public StructuralNavigation SearchNavigation
        {
            get { return (this._searchNavigation); }
            private set
            {
                if (this._searchNavigation != null)
                {
                    this._searchNavigation.ParentTemplate = null;
                }
                this._searchNavigation = value;
                if (this._searchNavigation != null)
                {
                    this._searchNavigation.ParentTemplate = this.ParentTemplate;
                }
            }
        }
        /// <summary>
        /// Declares whether the tree view has to be enabled at the site level or not, optional attribute.
        /// </summary>
        public Boolean EnableTreeView { get; set; }

        /// <summary>
        /// Declares whether the New Page ribbon command will automatically create a navigation item for the newly created page, optional attribute.
        /// </summary>
        public Boolean AddNewPagesToNavigation { get; set; }

        /// <summary>
        /// Declares whether the New Page ribbon command will automatically create a friendly URL for the newly created page, optional attribute.
        /// </summary>
        public Boolean CreateFriendlyUrlsForNewPages { get; set; }

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
        /// <param name="searchNavigation">SearchNavigation object</param>
        public Navigation(GlobalNavigation globalNavigation = null, CurrentNavigation currentNavigation = null, StructuralNavigation searchNavigation = null)
        {
            this.GlobalNavigation = globalNavigation;
            this.CurrentNavigation = currentNavigation;
            this.SearchNavigation = searchNavigation;
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|",
                (this.GlobalNavigation != null ? this.GlobalNavigation.GetHashCode() : 0),
                (this.CurrentNavigation != null ? this.CurrentNavigation.GetHashCode() : 0),
                (this.SearchNavigation != null ? this.SearchNavigation.GetHashCode() : 0),
                this.EnableTreeView.GetHashCode(),
                this.AddNewPagesToNavigation.GetHashCode(),
                this.CreateFriendlyUrlsForNewPages.GetHashCode()
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

            return (this.GlobalNavigation.Equals(other.GlobalNavigation) &&
                    this.CurrentNavigation.Equals(other.CurrentNavigation) &&
                    (this.SearchNavigation != null && other.SearchNavigation != null ? this.SearchNavigation.Equals(other.SearchNavigation) : this.SearchNavigation == null && other.SearchNavigation == null ? true : false) &&
                    this.EnableTreeView == other.EnableTreeView &&
                    this.AddNewPagesToNavigation == other.AddNewPagesToNavigation &&
                    this.CreateFriendlyUrlsForNewPages == other.CreateFriendlyUrlsForNewPages
                    );
        }

        #endregion
    }
}
