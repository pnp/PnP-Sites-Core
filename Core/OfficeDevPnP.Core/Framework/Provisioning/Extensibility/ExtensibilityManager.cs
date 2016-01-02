using System;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    ///  Provisioning Framework Component that is used for invoking custom providers during the provisioning process.
    /// </summary>
    public class ExtensibilityManager
    {
        private Dictionary<Provider, Object> providerCache = new Dictionary<Provider, Object>();

        /// <summary>
        /// Method to Invoke Custom Provisioning Token Providers which implement the IProvisioningExtensibilityTokenProvider interface.
        /// Ensure the ClientContext is not disposed in the custom provider.
        /// </summary>
        /// <param name="ctx">Authenticated ClientContext that is passed to the custom provider.</param>
        /// <param name="provider">A custom Extensibility Provisioning Provider</param>
        /// <param name="template">ProvisioningTemplate that is passed to the custom provider</param>
        /// <exception cref="ExtensiblityPipelineException"></exception>
        /// <exception cref="ArgumentException">Provider.Assembly or Provider.Type is NullOrWhiteSpace></exception>
        /// <exception cref="ArgumentNullException">ClientContext is Null></exception>
        public IEnumerable<TokenDefinition> ExecuteTokenProviderCallOut(ClientContext ctx, Provider provider, ProvisioningTemplate template)
        {
            var _loggingSource = "OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensibilityManager.ExecuteTokenProviderCallOut";

            if (ctx == null)
                throw new ArgumentNullException(CoreResources.Provisioning_Extensibility_Pipeline_ClientCtxNull);

            if (string.IsNullOrWhiteSpace(provider.Assembly))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);

            if (string.IsNullOrWhiteSpace(provider.Type))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_TypeName);

            try
            { 
                var _instance = GetProviderInstance(provider) as IProvisioningExtensibilityTokenProvider;
                if (_instance != null)
                {
                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_BeforeInvocation,
                        provider.Assembly,
                        provider.Type);

                    var tokens = _instance.GetTokens(ctx, template, provider.Configuration);

                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_Success,
                        provider.Assembly,
                        provider.Type);
                
                    return tokens;
                }
                return new List<TokenDefinition>();
                }
            catch (Exception ex)
            {
                string _message = string.Format(
                    CoreResources.Provisioning_Extensibility_Pipeline_Exception,
                    provider.Assembly,
                    provider.Type,
                    ex);
                Log.Error(_loggingSource, _message);
                throw new ExtensiblityPipelineException(_message, ex);
            }      
        }


        /// <summary>
        /// Method to Invoke Custom Provisioning Providers. 
        /// Ensure the ClientContext is not disposed in the custom provider.
        /// </summary>
        /// <param name="ctx">Authenticated ClientContext that is passed to the custom provider.</param>
        /// <param name="provider">A custom Extensibility Provisioning Provider</param>
        /// <param name="template">ProvisioningTemplate that is passed to the custom provider</param>
        /// <exception cref="ExtensiblityPipelineException"></exception>
        /// <exception cref="ArgumentException">Provider.Assembly or Provider.Type is NullOrWhiteSpace></exception>
        /// <exception cref="ArgumentNullException">ClientContext is Null></exception>
        public void ExecuteExtensibilityCallOut(ClientContext ctx, Provider provider, ProvisioningTemplate template)
        {
            var _loggingSource = "OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensibilityManager.ExecuteCallout";

            if (ctx == null)
                throw new ArgumentNullException(CoreResources.Provisioning_Extensibility_Pipeline_ClientCtxNull);

            if (string.IsNullOrWhiteSpace(provider.Assembly))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);

            if (string.IsNullOrWhiteSpace(provider.Type))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_TypeName);

            try
            {
                var _instance = GetProviderInstance(provider) as IProvisioningExtensibilityProvider;
                if (_instance != null)
                { 
                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_BeforeInvocation,
                        provider.Assembly,
                        provider.Type);

                    _instance.ProcessRequest(ctx, template, provider.Configuration);

                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_Success,
                        provider.Assembly,
                        provider.Type);
                }
            }
            catch(Exception ex)
            {
                string _message = string.Format(
                    CoreResources.Provisioning_Extensibility_Pipeline_Exception, 
                    provider.Assembly, 
                    provider.Type, 
                    ex);
                Log.Error(_loggingSource, _message);
                throw new ExtensiblityPipelineException(_message, ex);
             
            }
        }

        private object GetProviderInstance(Provider provider)
        {
            if (string.IsNullOrWhiteSpace(provider.Assembly))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);

            if (string.IsNullOrWhiteSpace(provider.Type))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_TypeName);

            if (!providerCache.ContainsKey(provider))
            {
                var _instance = Activator.CreateInstance(provider.Assembly, provider.Type).Unwrap();
                providerCache.Add(provider, _instance);
            }
            return providerCache[provider];
        }    
    }
}
