using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Base abstract class for Team and TeamTemplate types
    /// </summary>
    public abstract partial class BaseTeam: BaseModel, IEquatable<BaseTeam>
    {
        #region Public Members

        /// <summary>
        /// The Display Name of the Team
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// The Description of the Team
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The Classification for the Team
        /// </summary>
        public String Classification { get; set; }

        /// <summary>
        /// The Visibility for the Team
        /// </summary>
        public TeamVisibility Visibility { get; set; }

        /// <summary>
        /// The Photo for the Team
        /// </summary>
        public String Photo { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                DisplayName.GetHashCode(),
                Description.GetHashCode(),
                Classification?.GetHashCode() ?? 0,
                Visibility.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with BaseTeam class
        /// </summary>
        /// <param name="obj">Object that represents BaseTeam</param>
        /// <returns>Checks whether object is BaseTeam class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is BaseTeam))
            {
                return (false);
            }
            return (Equals((BaseTeam)obj));
        }

        /// <summary>
        /// Compares BaseTeam object based on DisplayName, Description, Classification, and Visibility
        /// </summary>
        /// <param name="other">BaseTeam Class object</param>
        /// <returns>true if the BaseTeam object is equal to the current object; otherwise, false.</returns>
        public bool Equals(BaseTeam other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.DisplayName == other.DisplayName &&
                this.Description == other.Description &&
                this.Classification == other.Classification &&
                this.Visibility == other.Visibility
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the Visibility for a Microsoft Team
    /// </summary>
    public enum TeamVisibility
    {
        /// <summary>
        /// Defines a Private Team
        /// </summary>
        Private,
        /// <summary>
        /// Defines a Public Team
        /// </summary>
        Public
    }
}
