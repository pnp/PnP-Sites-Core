using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// Holds PnP Core library identification tag and user-agent, and a tool to get tenant administration url based the URL of the web
    /// </summary>
    public static class PnPCoreUtilities
    {
        /// <summary>
        /// Get's a tag that identifies the PnP Core library
        /// </summary>
        /// <returns>PnP Core library identification tag</returns>
        public static string PnPCoreVersionTag
        {
            get
            {
                return (PnPCoreVersionTagLazy.Value);
            }
        }

        private static Lazy<String> PnPCoreVersionTagLazy = new Lazy<String>(
            () => {
                Assembly coreAssembly = Assembly.GetExecutingAssembly();
                String result = $"PnPCore:{((AssemblyFileVersionAttribute) coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version.Split('.')[2]}";
                return (result);
            }, 
            true);

        /// <summary>
        /// Get's a tag that identifies the PnP Core library for UserAgent string
        /// </summary>
        /// <returns>PnP Core library identification user-agent</returns>
        public static string PnPCoreUserAgent
        {
            get
            {
                return (PnPCoreUserAgentLazy.Value);
            }
        }

        private static Lazy<String> PnPCoreUserAgentLazy = new Lazy<String>(
            () => {
                Assembly coreAssembly = Assembly.GetExecutingAssembly();         
                String result = $"NONISV|SharePointPnP|PnPCore/{((AssemblyFileVersionAttribute)coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version}";
                return (result);
            },
            true);

        /// <summary>
        /// Returns the tenant administration url based upon the URL of the web
        /// </summary>
        /// <param name="web"></param>
        /// <returns></returns>
        public static string GetTenantAdministrationUrl(this Web web)
        {
            var url = web.EnsureProperty(w => w.Url);
            var uri = new Uri(url);
            var uriParts = uri.Host.Split('.');
            if (uriParts[0].EndsWith("-admin")) return uri.OriginalString;
            if (!uriParts[0].EndsWith("-admin"))
                return $"https://{uriParts[0]}-admin.{string.Join(".", uriParts.Skip(1))}";
            return null;
        }

#if !SP2013
        /// <summary>
        /// Executes the specified method with Return Value Cache disabled on the Client Runtime Context.
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="context"></param>
        /// <param name="disabledReturnValueCacheCode">The code that is to run with Return Value Cache disabled.</param>
        /// <returns>Returns value from the code specified by <paramref name="disabledReturnValueCacheCode"/></returns>
        public static TResult RunWithDisableReturnValueCache<TResult>(ClientRuntimeContext context, Func<TResult> disabledReturnValueCacheCode)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (disabledReturnValueCacheCode == null)
            {
                throw new ArgumentNullException(nameof(disabledReturnValueCacheCode));
            }

            bool disableReturnValueCache = context.DisableReturnValueCache;

            try
            {
                context.DisableReturnValueCache = true;
                return disabledReturnValueCacheCode();
            }
            catch
            {
                throw;
            }
            finally
            {
                context.DisableReturnValueCache = disableReturnValueCache;
            }
        }

        /// <summary>
        /// Executes the specified method with Return Value Cache disabled on the Client Runtime Context.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="disabledReturnValueCacheCode">The code that is to run with Return Value Cache disabled.</param>
        public static void RunWithDisableReturnValueCache(ClientRuntimeContext context, Action disabledReturnValueCacheCode)
        {
            RunWithDisableReturnValueCache<object>(context, () => { disabledReturnValueCacheCode(); return default; } );
        }
#endif
    }
}
