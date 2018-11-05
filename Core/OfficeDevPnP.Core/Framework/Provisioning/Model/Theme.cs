using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Theme to provision at tenant level
    /// </summary>
    public partial class Theme : BaseModel, IEquatable<Theme>
    {
        #region Public Members

        /// <summary>
        /// Defines the Name of the Theme
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Declares if the Theme Is Inverted
        /// </summary>
        public Boolean IsInverted { get; set; }

        /// <summary>
        /// Defines the Palette of the Theme
        /// </summary>
        /// <remarks>
        /// It has to be a JSON object representing a dictionary of colors
        /// </remarks>
        public String Palette { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}",
                Name?.GetHashCode() ?? 0,
                IsInverted.GetHashCode(),
                Palette?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Theme class
        /// </summary>
        /// <param name="obj">Object that represents Theme</param>
        /// <returns>Checks whether object is Theme class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Theme))
            {
                return (false);
            }
            return (Equals((Theme)obj));
        }

        /// <summary>
        /// Compares Theme object based on Name, IsInverted, and Palette
        /// </summary>
        /// <param name="other">Theme Class object</param>
        /// <returns>true if the Theme object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Theme other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                this.IsInverted == other.IsInverted &&
                this.Palette == other.Palette
                );
        }

        #endregion
    }
}
