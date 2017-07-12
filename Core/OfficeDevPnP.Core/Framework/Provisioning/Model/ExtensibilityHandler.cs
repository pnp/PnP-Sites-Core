using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object for Extensiblity Call out
    /// </summary>
    public class ExtensibilityHandler : BaseModel, IEquatable<ExtensibilityHandler>
    {
        #region Properties
        /// <summary>
        /// Gets or sets Enabled property for Extensibility handling.
        /// </summary>
        public bool Enabled
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets Assembly property for Extensibility handling.
        /// </summary>
        public string Assembly
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets Type property for Extensibility handling.
        /// </summary>
        public string Type
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets Configuration property for Extensibility handling.
        /// </summary>
        public string Configuration { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                (this.Assembly != null ? this.Assembly.GetHashCode() : 0),
                (this.Configuration != null ? this.Configuration.GetHashCode() : 0),
                this.Enabled.GetHashCode(),
                (this.Type != null ? this.Type.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ExtensibilityHandler
        /// </summary>
        /// <param name="obj">Object that represents ExtensibilityHandler</param>
        /// <returns>true if the current object is equal to the ExtensibilityHandler</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ExtensibilityHandler))
            {
                return (false);
            }
            return (Equals((ExtensibilityHandler)obj));
        }

        /// <summary>
        /// Compares ExtensibilityHandler object based on Assembly, Configuration, Enabled and Type properties.
        /// </summary>
        /// <param name="other">ExtensibilityHandler object</param>
        /// <returns>true if the ExtensibilityHandler object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ExtensibilityHandler other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Assembly == other.Assembly &&
                this.Configuration == other.Configuration &&
                this.Enabled == other.Enabled &&
                this.Type == other.Type);
        }

        #endregion
    }
}
