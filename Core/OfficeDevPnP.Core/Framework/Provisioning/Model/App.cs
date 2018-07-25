using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines an Add-in to provision
    /// </summary>
    public partial class App : BaseModel, IEquatable<App>
    {
        #region Public Members

        /// <summary>
        /// Defines the AppId for the App to manage
        /// </summary>
        public String AppId { get; set; }

        /// <summary>
        /// Defines the Action for the App to manage
        /// </summary>
        /// <remarks>
        /// Possible values are: Install, Update, Uninstall.
        /// </remarks>
        public AppAction Action { get; set; }

        /// <summary>
        /// Defines whether the package will be handled synchronously or asynchronously
        /// </summary>
        /// <remarks>
        /// Possible values are: Synchronously, Asynchronously.
        /// </remarks>
        public SyncMode SyncMode { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                this.AppId?.GetHashCode() ?? 0,
                this.Action.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with App class
        /// </summary>
        /// <param name="obj">Object that represents App</param>
        /// <returns>Checks whether object is App class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is App))
            {
                return (false);
            }
            return (Equals((App)obj));
        }

        /// <summary>
        /// Compares App object based on PackagePath and source
        /// </summary>
        /// <param name="other">App Class object</param>
        /// <returns>true if the App object is equal to the current object; otherwise, false.</returns>
        public bool Equals(App other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AppId == other.AppId &&
                this.Action == other.Action
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the action to execute against the package
    /// </summary>
    public enum AppAction
    {
        /// <summary>
        /// Instructs the engine to install the package in the site collection.
        /// </summary>
        Install,
        /// <summary>
        /// Instructs the engine to update the package in the site collection.
        /// </summary>
        Update,
        /// <summary>
        /// Instructs the engine to uninstall the package from the site collection.
        /// </summary>
        Uninstall,
    }

    /// <summary>
    /// Defines whether the package will be handled synchronously or asynchronously
    /// </summary>
    public enum SyncMode
    {
        /// <summary>
        /// Defines that the package will be handled synchronously.
        /// </summary>
        Synchronously,
        /// <summary>
        /// Defines that the package will be handled asynchronously.
        /// </summary>
        Asynchronously,
    }
}
