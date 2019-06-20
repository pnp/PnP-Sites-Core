using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that defines a Composed Look in the Provision Template
    /// </summary>
    public partial class ComposedLook : BaseModel, IEquatable<ComposedLook>
    {
        #region Constructors

        /// <summary>
        /// Constructor for ComposedLook class
        /// </summary>
        static ComposedLook()
        {
            Empty = new ComposedLook();
        }

        /// <summary>
        /// Constructor for ComposedLook class
        /// </summary>
        public ComposedLook()
        {
        }

        #endregion

        #region Private members

        private static ComposedLook _empty;

        public static ComposedLook Empty
        {
            private set { _empty = value; }
            get { return (_empty); }
        }

        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the Name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the ColorFile
        /// </summary>
        public string ColorFile { get; set; }

        /// <summary>
        /// Gets or sets the FontFile
        /// </summary>
        public string FontFile { get; set; }

        /// <summary>
        /// Gets or sets the Background Image 
        /// </summary>
        public string BackgroundFile { get; set; }

        /// <summary>
        /// Gets or sets the Version of the ComposedLook.
        /// </summary>
        public int Version { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}",
                BackgroundFile?.GetHashCode() ?? 0,
                ColorFile?.GetHashCode() ?? 0,
                FontFile?.GetHashCode() ?? 0,
                Name?.GetHashCode() ?? 0,
                this.Version.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ComposedLook
        /// </summary>
        /// <param name="obj">Object that represents ComposedLook</param>
        /// <returns>true if the current object is equal to the ComposedLook</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ComposedLook))
            {
                return(false);
            }
            return (Equals((ComposedLook)obj));
        }

        /// <summary>
        /// Compares ComposedLook object based on BackgroundFile, ColorFile, FontFile, Name and Version.
        /// </summary>
        /// <param name="other">ComposedLook object</param>
        /// <returns>true if the ComposedLook object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ComposedLook other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.BackgroundFile == other.BackgroundFile &&
                this.ColorFile == other.ColorFile &&
                this.FontFile == other.FontFile &&
                this.Name == other.Name &&
                this.Version == other.Version);
        }

        #endregion
    }
}
