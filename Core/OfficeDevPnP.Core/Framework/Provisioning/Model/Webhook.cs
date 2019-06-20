using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Webhook
    /// </summary>
    public partial class Webhook : BaseModel, IEquatable<Webhook>
    {
        #region Public Members

        /// <summary>
        /// Defines the Server Notification URL of the Webhook, required attribute.
        /// </summary>
        public String ServerNotificationUrl { get; set; }

        /// <summary>
        /// Defines the expire days for the subscription of the Webhook, required attribute.
        /// </summary>
        /// <remarks>
        /// The maximum value is 6 months (i.e. 180 days)
        /// </remarks>
        public Int32 ExpiresInDays { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                ServerNotificationUrl?.GetHashCode() ?? 0,
                ExpiresInDays.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Webhook class
        /// </summary>
        /// <param name="obj">Object that represents Webhook</param>
        /// <returns>Checks whether object is Webhook class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Webhook))
            {
                return (false);
            }
            return (Equals((Webhook)obj));
        }

        /// <summary>
        /// Compares Webhook object based on ServerNotificationUrl and ExpiresInDays
        /// </summary>
        /// <param name="other">Webhook Class object</param>
        /// <returns>true if the Webhook object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Webhook other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ServerNotificationUrl == other.ServerNotificationUrl &&
                this.ExpiresInDays == other.ExpiresInDays
                );
        }

        #endregion
    }
}
