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

        #endregion

        #region Constructors

        public Navigation()
        {

        }

        public Navigation(GlobalNavigation globalNavigation = null, CurrentNavigation currentNavigation = null)
        {
            this.GlobalNavigation = globalNavigation;
            this.CurrentNavigation = currentNavigation;
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                (this.GlobalNavigation != null ? this.GlobalNavigation.GetHashCode() : 0),
                (this.CurrentNavigation != null ? this.CurrentNavigation.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Navigation))
            {
                return (false);
            }
            return (Equals((Navigation)obj));
        }

        public bool Equals(Navigation other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.GlobalNavigation == other.GlobalNavigation &&
                this.CurrentNavigation == other.CurrentNavigation
                );
        }

        #endregion
    }
}
