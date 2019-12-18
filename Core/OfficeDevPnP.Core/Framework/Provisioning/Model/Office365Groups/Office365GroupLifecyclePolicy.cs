using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Office365Groups
{
    /// <summary>
    /// Defines a Office365GroupLifecyclePolicy object
    /// </summary>
    public partial class Office365GroupLifecyclePolicy : BaseModel, IEquatable<Office365GroupLifecyclePolicy>
    {
        #region Public members

        /// <summary>
        /// The ID of the target Office365GroupLifecyclePolicy
        /// </summary>
        public String ID { get; set; }

        /// <summary>
        /// The GroupLifetimeInDays of the target Office365GroupLifecyclePolicy
        /// </summary>
        public Int32 GroupLifetimeInDays { get; set; }

        /// <summary>
        /// The AlternateNotificationEmails of the target Office365GroupLifecyclePolicy
        /// </summary>
        public String AlternateNotificationEmails { get; set; }

        /// <summary>
        /// The AlternateNotificationEmails of the target Office365GroupLifecyclePolicy
        /// </summary>
        public ManagedGroupTypes ManagedGroupTypes { get; set; }

        #endregion

        #region Constructors

        public Office365GroupLifecyclePolicy() : base()
        {
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                ID?.GetHashCode() ?? 0,
                GroupLifetimeInDays.GetHashCode(),
                AlternateNotificationEmails?.GetHashCode() ?? 0,
                ManagedGroupTypes.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Office365GroupLifecyclePolicy class
        /// </summary>
        /// <param name="obj">Object that represents Office365GroupLifecyclePolicy</param>
        /// <returns>Checks whether object is Office365GroupLifecyclePolicy class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Office365GroupLifecyclePolicy))
            {
                return (false);
            }
            return (Equals((Office365GroupLifecyclePolicy)obj));
        }

        /// <summary>
        /// Compares Office365GroupLifecyclePolicy object based on ID, GroupLifetimeInDays, AlternateNotificationEmails,
        /// and ManagedGroupTypes
        /// </summary>
        /// <param name="other">User Office365GroupLifecyclePolicy object</param>
        /// <returns>true if the Office365GroupLifecyclePolicy object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Office365GroupLifecyclePolicy other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ID == other.ID &&
                this.GroupLifetimeInDays == other.GroupLifetimeInDays &&
                this.AlternateNotificationEmails == other.AlternateNotificationEmails &&
                this.ManagedGroupTypes == other.ManagedGroupTypes
                );
        }

        #endregion
    }

    public enum ManagedGroupTypes
    {
        /// <summary>
        /// All the Managed Group Types
        /// </summary>
        All,
        /// <summary>
        /// Selected Managed Group Types
        /// </summary>
        Selected,
        /// <summary>
        /// None of the Managed Group Types
        /// </summary>
        None,
    }
}
