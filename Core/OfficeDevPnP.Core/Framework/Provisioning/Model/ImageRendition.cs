using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines an available Image Rendition for the current Publishing site.
    /// </summary>
    public partial class ImageRendition : BaseModel, IEquatable<ImageRendition>
    {
        #region Public Members

        /// <summary>
        /// Defines the name of the Image Rendition for the current Publishing site, required attribute.
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Defines the width of the Image Rendition for the current Publishing site, required attribute.
        /// </summary>
        public Int32 Width { get; set; }

        /// <summary>
        /// Defines the height of the Image Rendition for the current Publishing site, required attribute.
        /// </summary>
        public Int32 Height { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.Name?.GetHashCode() ?? 0,
                this.Width.GetHashCode(),
                this.Height.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ImageRendition class
        /// </summary>
        /// <param name="obj">Object that represents ImageRendition</param>
        /// <returns>Checks whether object is ImageRendition class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ImageRendition))
            {
                return (false);
            }
            return (Equals((ImageRendition)obj));
        }

        /// <summary>
        /// Compares ImageRendition object based on Name, Width, and Height
        /// </summary>
        /// <param name="other">ImageRendition Class object</param>
        /// <returns>true if the ImageRendition object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ImageRendition other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name &&
                this.Width == other.Width &&
                this.Height == other.Height
                );
        }

        #endregion
    }
}
