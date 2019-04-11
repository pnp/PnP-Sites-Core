using System;
using System.Collections.Generic;
using System.Linq;
#if !NETSTANDARD2_0
using System.Runtime.Remoting.Messaging;
#endif
using System.Text;
using System.Threading.Tasks;
#if NETSTANDARD2_0
using System.Threading;
#endif

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Asynchronous delegate to acquire an Access Token to get access to a target resource
    /// </summary>
    /// <param name="resource">The Resource to access</param>
    /// <param name="scope">The required Permission Scope</param>
    /// <returns>The Access Token to access the target resource</returns>
    public delegate Task<String> AcquireTokenAsyncDelegate(String resource, String scope);

    /// <summary>
    /// Asynchronous delegate to get a cookie to access a target resource
    /// </summary>
    /// <param name="resource">The Resource to access</param>
    /// <returns>The Cookie to access the target resource</returns>
    public delegate Task<String> AcquireCookieAsyncDelegate(String resource);

    /// <summary>
    /// Class to wrap any PnP Provisioning process in order to share the same security context across different Object Handlers
    /// </summary>
    public class PnPProvisioningContext : IDisposable
    {
        private readonly PnPProvisioningContext _previous;

        public AcquireTokenAsyncDelegate AcquireTokenAsync { get; private set; }
        public AcquireCookieAsyncDelegate AcquireCookieAsync { get; private set; }

        public Dictionary<String, Object> Properties { get; private set; } = 
            new Dictionary<string, object>();

        public PnPProvisioningContext(
            AcquireTokenAsyncDelegate acquireTokenAsyncDelegate = null,
            AcquireCookieAsyncDelegate acquireCookieAsyncDelegate = null,
            Dictionary<String, Object> properties = null)
        {
            // Save the delegate to acquire the access token
            this.AcquireTokenAsync = acquireTokenAsyncDelegate;

            // Save the delegate to acquire the cookie
            this.AcquireCookieAsync = acquireCookieAsyncDelegate;

            // Add the initial set of properties, if any
            if (properties != null)
            {
                foreach (var p in properties)
                {
                    this.Properties.Add(p.Key, p.Value);
                }
            }

            // Save the previous context, if any
            this._previous = Current;

            // Set the new context to this
            Current = this;
        }

        public String AcquireToken(String resource, String scope)
        {
            return(this.AcquireTokenAsync(resource, scope).GetAwaiter().GetResult());
        }

        public String AcquireCookie(String resource)
        {
            return (this.AcquireCookieAsync(resource).GetAwaiter().GetResult());
        }

        ~PnPProvisioningContext()
        {
            Dispose(false);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                Current = this._previous;
            }
        }

#if !NETSTANDARD2_0
        public static PnPProvisioningContext Current
        {
            get { return CallContext.LogicalGetData(nameof(PnPProvisioningContext)) as PnPProvisioningContext; }
            set
            {
                System.Configuration.ConfigurationManager.GetSection("system.xml/xmlReader");
                CallContext.LogicalSetData(nameof(PnPProvisioningContext), value);
            }
        }
#else
        private static AsyncLocal<PnPProvisioningContext> _pnpSerializationScope = new AsyncLocal<PnPProvisioningContext>();

        public static PnPProvisioningContext Current
        {
            get { return _pnpSerializationScope.Value; }
            set
            {
                _pnpSerializationScope.Value = value;
            }
        }
#endif
    }
}
