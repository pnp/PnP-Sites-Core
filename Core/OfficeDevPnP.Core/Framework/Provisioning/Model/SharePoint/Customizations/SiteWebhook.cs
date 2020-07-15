using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a SiteWebhook to provision
    /// </summary>
    public partial class SiteWebhook : Webhook
    {        
        #region Public Members

        /// <summary>
        /// Defines the Server Notification URL of the Webhook, required attribute.
        /// </summary>
        public SiteWebhookType SiteWebhookType { get; set; }

        // CS0108 - hides inherited member.
        /*
        /// <summary>
        /// Defines the expire days for the subscription of the Webhook, required attribute.
        /// </summary>
        /// <remarks>
        /// The maximum value is 6 months (i.e. 180 days)
        /// </remarks>
        public Int32 ExpiresInDays { get; set; }
        */

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                ServerNotificationUrl?.GetHashCode() ?? 0,
                ExpiresInDays.GetHashCode(),
                SiteWebhookType.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteWebhook class
        /// </summary>
        /// <param name="obj">Object that represents SiteWebhook</param>
        /// <returns>Checks whether object is SiteWebhook class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteWebhook))
            {
                return (false);
            }
            return (Equals((SiteWebhook)obj));
        }

        /// <summary>
        /// Compares SiteWebhook object based on ServerNotificationUrl, ExpiresInDays, and SiteWebhookType
        /// </summary>
        /// <param name="other">SiteWebhook Class object</param>
        /// <returns>true if the SiteWebhook object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteWebhook other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ServerNotificationUrl == other.ServerNotificationUrl &&
                this.ExpiresInDays == other.ExpiresInDays &&
                this.SiteWebhookType == other.SiteWebhookType
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the various flavors of a SiteWebhook
    /// </summary>
    public enum SiteWebhookType
    {
        /// <summary>
        /// A WebHook for the WebCreated event of Site.
        /// </summary>
        WebCreated,
        /// <summary>
        /// A WebHook for the WebMoved event of Site.
        /// </summary>
        WebMoved,
        /// <summary>
        /// A WebHook for the WebDeleted event of Site.
        /// </summary>
        WebDeleted,
        /// <summary>
        /// A WebHook for the ListAdded event of Site.
        /// </summary>
        ListAdded,
        /// <summary>
        /// A WebHook for the ListCreated event of Site.
        /// </summary>
        ListCreated,

    }
}
