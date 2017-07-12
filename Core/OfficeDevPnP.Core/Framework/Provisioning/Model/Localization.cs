using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object used in the Provisioning template that defines a Localization item
    /// </summary>
    public partial class Localization : BaseModel, IEquatable<Localization>
    {
        #region Properties

        /// <summary>
        /// The Locale ID of a Localization Language
        /// </summary>
        public Int32 LCID { get; set; }

        /// <summary>
        /// The Name of a Localization Language
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// The path to the .RESX (XML) resource file for the current Localization
        /// </summary>
        public String ResourceFile { get; set; }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for Localization class
        /// </summary>
        public Localization() { }

        /// <summary>
        /// Constructor for Localization class
        /// </summary>
        /// <param name="lcid">Language id for localization</param>
        /// <param name="name">Name for localization</param>
        /// <param name="resourceFile">Resource file for localization</param>
        public Localization(Int32 lcid, String name, String resourceFile)
        {
            this.LCID = lcid;
            this.Name = name;
            this.ResourceFile = resourceFile;
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
                (this.LCID.GetHashCode()),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.ResourceFile != null ? this.ResourceFile.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Localization
        /// </summary>
        /// <param name="obj">Object that represents Localization</param>
        /// <returns>true if the current object is equal to the Localization</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Localization))
            {
                return (false);
            }
            return (Equals((Localization)obj));
        }

        /// <summary>
        /// Compares Localization object based on LCID, Name and ResourceFile
        /// </summary>
        /// <param name="other">Localization object</param>
        /// <returns>true if the Localization object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Localization other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.LCID == other.LCID &&
                    this.Name == other.Name &&
                    this.ResourceFile == other.ResourceFile 
                );

        }

        #endregion
    }
}
