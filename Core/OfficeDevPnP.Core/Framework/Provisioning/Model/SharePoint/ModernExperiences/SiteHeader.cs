using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Header settings for the target site
    /// </summary>
    public partial class SiteHeader : BaseModel, IEquatable<SiteHeader>
    {
        #region Public Members

        /// <summary>
        /// Defines the Layout of the Header
        /// </summary>
        public SiteHeaderLayout Layout { get; set; }

        /// <summary>
        /// Defines the Menu Style
        /// </summary>
        public SiteHeaderMenuStyle MenuStyle { get; set; }

        /// <summary>
        /// Defines the Background Emphasis of the Header
        /// </summary>
        public Emphasis BackgroundEmphasis { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                Layout.GetHashCode(),
                MenuStyle.GetHashCode(),
                BackgroundEmphasis.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteHeader class
        /// </summary>
        /// <param name="obj">Object that represents SiteHeader</param>
        /// <returns>Checks whether object is SiteHeader class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteHeader))
            {
                return (false);
            }
            return (Equals((SiteHeader)obj));
        }

        /// <summary>
        /// Compares SiteHeader object based on Layout, MenuStyle, and BackgroundEmphasis
        /// </summary>
        /// <param name="other">SiteHeader Class object</param>
        /// <returns>true if the SiteHeader object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteHeader other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Layout == other.Layout &&
                this.MenuStyle == other.MenuStyle && 
                this.BackgroundEmphasis == other.BackgroundEmphasis
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the Layouts available for the Header of a Site
    /// </summary>
    public enum SiteHeaderLayout
    {
        /// <summary>
        /// Defines the Standard Layout for the Site Header
        /// </summary>
        Standard,
        /// <summary>
        /// Defines the Compact Layout for the Site Header
        /// </summary>
        Compact,
    }

    /// <summary>
    /// Defines the Menu Styles available for the Header of a Site
    /// </summary>
    public enum SiteHeaderMenuStyle
    {
        /// <summary>
        /// Defines the MegaMenu Style for the Site Menu
        /// </summary>
        MegaMenu,
        /// <summary>
        /// Defines the Cascading Style for the Site Menu
        /// </summary>
        Cascading,
    }

    /// <summary>
    /// Defines the Emphasis for a modern section
    /// </summary>
    public enum Emphasis
    {
        /// <summary>
        /// Defines No emphasis for the Site Header.
        /// </summary>
        None,
        /// <summary>
        /// Defines the Neutral emphasis for the Site Header
        /// </summary>
        Neutral,
        /// <summary>
        /// Defines the Soft emphasis for the Site Header.
        /// </summary>
        Soft,
        /// <summary>
        /// Defines the Strong emphasis for the Site Header.
        /// </summary>
        Strong
    }
}
