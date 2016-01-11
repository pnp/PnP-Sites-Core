using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that defines a Composed Look in the Provision Template
    /// </summary>
    public partial class ComposedLook : BaseModel, IEquatable<ComposedLook>
    {
        #region Constructors

        static ComposedLook()
        {
            Empty = new ComposedLook();
        }

        public ComposedLook()
        {
        }

        #endregion 

        private static ComposedLook _empty;

        public static ComposedLook Empty
        {
            private set { _empty = value; }
            get { return (_empty); }
        }

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

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|",
                (this.BackgroundFile != null ? this.BackgroundFile.GetHashCode() : 0),
                (this.ColorFile != null ? this.ColorFile.GetHashCode() : 0),
                (this.FontFile != null ? this.FontFile.GetHashCode() : 0),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                this.Version.GetHashCode()
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ComposedLook))
            {
                return(false);
            }
            return (Equals((ComposedLook)obj));
        }

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
