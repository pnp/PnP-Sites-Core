using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public abstract partial class ProvisioningWebhookBase : BaseModel, IEquatable<ProvisioningWebhookBase>
    {
        #region Public Members

        /// <summary>
        /// Defines the custom parameters for the Provisioning Template Webhook
        /// </summary>
        public Dictionary<String, String> Parameters { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Defines the .app file of the SharePoint Add-in to provision
        /// </summary>
        public ProvisioningTemplateWebhookKind Kind { get; set; }

        /// <summary>
        /// Defines the URL of a Provisioning Template Webhook, can be a replaceable string
        /// </summary>
        public String Url { get; set; }

        /// <summary>
        /// Defines how to call the target Webhook URL
        /// </summary>
        public ProvisioningTemplateWebhookMethod Method { get; set; }

        /// <summary>
        /// Defines how to format the request body for HTTP POST requests
        /// </summary>
        public ProvisioningTemplateWebhookBodyFormat BodyFormat { get; set; }

        /// <summary>
        /// Defines whether the Provisioning Template Webhook should be executed asychronously or not
        /// </summary>
        public Boolean Async { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|",
                this.Parameters.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                Kind.GetHashCode(),
                Url?.GetHashCode() ?? 0,
                Method.GetHashCode(),
                BodyFormat.GetHashCode(),
                Async.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ProvisioningWebhookBase class
        /// </summary>
        /// <param name="obj">Object that represents ProvisioningWebhookBase</param>
        /// <returns>Checks whether object is ProvisioningWebhookBase class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningWebhookBase))
            {
                return (false);
            }
            return (Equals((ProvisioningWebhookBase)obj));
        }

        /// <summary>
        /// Compares ProvisioningWebhookBase object based on PackagePath and source
        /// </summary>
        /// <param name="other">ProvisioningWebhookBase Class object</param>
        /// <returns>true if the ProvisioningWebhookBase object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ProvisioningWebhookBase other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Parameters.DeepEquals(other.Parameters) && 
                this.Kind == other.Kind &&
                this.Url == other.Url &&
                this.Method == other.Method &&
                this.BodyFormat == other.BodyFormat &&
                this.Async == other.Async
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the kind of a Provisioning Template Webhook
    /// </summary>
    public enum ProvisioningTemplateWebhookKind
    {
        /// <summary>
        /// Provisioning Started
        /// </summary>
        ProvisioningStarted,
        /// <summary>
        /// Object Handler Provisioning Started
        /// </summary>
        ObjectHandlerProvisioningStarted,
        /// <summary>
        /// Object Handler Provisioning Completed
        /// </summary>
        ObjectHandlerProvisioningCompleted,
        /// <summary>
        /// Provisioning Completed
        /// </summary>
        ProvisioningCompleted,
        /// <summary>
        /// An Exception Occurred
        /// </summary>
        ExceptionOccurred,
        /// <summary>
        /// Provisioning Template Started
        /// </summary>
        ProvisioningTemplateStarted,
        /// <summary>
        /// Provisioning Template Completed
        /// </summary>
        ProvisioningTemplateCompleted,
        /// <summary>
        /// Provisioning Exception Occurred
        /// </summary>
        ProvisioningExceptionOccurred,
    }

    /// <summary>
    /// Defines how to call the target Webhook URL
    /// </summary>
    public enum ProvisioningTemplateWebhookMethod
    {
        /// <summary>
        /// Invoke the Webhook with a HTTP GET request. Any Parameter optional will be in the querystring.
        /// </summary>
        GET,
        /// <summary>
        /// Invoke the Webhook with a HTTP POST request. Any Parameter optional will be in the request body.
        /// </summary>
        POST,
    }

    /// <summary>
    /// Defines how to format the request body for HTTP POST requests
    /// </summary>
    public enum ProvisioningTemplateWebhookBodyFormat
    {
        /// <summary>
        /// JSON format
        /// </summary>
        Json,
        /// <summary>
        /// XML format
        /// </summary>
        Xml,
        /// <summary>
        /// x-www-form-urlencoded format
        /// </summary>
        FormUrlEncoded,
    }
}
