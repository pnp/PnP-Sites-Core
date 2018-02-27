using System;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    ///  Provisioning Framework Component that is used for invoking custom providers during the provisioning process.
    /// </summary>
    public partial class ExtensibilityManager
    {
        private Dictionary<ExtensibilityHandler, Object> handlerCache = new Dictionary<ExtensibilityHandler, Object>();

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
        public IEnumerable<TokenDefinition> ExecuteTokenProviderCallOut(ClientContext ctx, ExtensibilityHandler provider, ProvisioningTemplate template)
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
        /// Method to Invoke Custom Provisioning Handlers. 
        /// </summary>
        /// <remarks>
        /// Ensure the ClientContext is not disposed in the custom provider.
        /// </remarks>
        /// <param name="ctx">Authenticated ClientContext that is passed to the custom provider.</param>
        /// <param name="handler">A custom Extensibility Provisioning Provider</param>
        /// <param name="template">ProvisioningTemplate that is passed to the custom provider</param>
        /// <param name="applyingInformation">The Provisioning Template application information object</param>
        /// <param name="tokenParser">The Token Parser used by the engine during template provisioning</param>
        /// <param name="scope">The PnPMonitoredScope of the current step in the pipeline</param>
        /// <exception cref="ExtensiblityPipelineException"></exception>
        /// <exception cref="ArgumentException">Provider.Assembly or Provider.Type is NullOrWhiteSpace></exception>
        /// <exception cref="ArgumentNullException">ClientContext is Null></exception>
        public void ExecuteExtensibilityProvisionCallOut(ClientContext ctx, ExtensibilityHandler handler, 
            ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation, 
            TokenParser tokenParser, PnPMonitoredScope scope)
        {
            var _loggingSource = "OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensibilityManager.ExecuteCallout";

            if (ctx == null)
                throw new ArgumentNullException(CoreResources.Provisioning_Extensibility_Pipeline_ClientCtxNull);

            if (string.IsNullOrWhiteSpace(handler.Assembly))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);

            if (string.IsNullOrWhiteSpace(handler.Type))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_TypeName);

            try
            {

                var _instance = GetProviderInstance(handler);
#pragma warning disable 618
                if (_instance is IProvisioningExtensibilityProvider)
#pragma warning restore 618
                {
                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_BeforeInvocation,
                        handler.Assembly,
                        handler.Type);

#pragma warning disable 618
                    (_instance as IProvisioningExtensibilityProvider).ProcessRequest(ctx, template, handler.Configuration);
#pragma warning restore 618

                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_Success,
                        handler.Assembly,
                        handler.Type);
                }
                else if (_instance is IProvisioningExtensibilityHandler)
                {
                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_BeforeInvocation,
                        handler.Assembly,
                        handler.Type);

                    (_instance as IProvisioningExtensibilityHandler).Provision(ctx, template, applyingInformation, tokenParser, scope, handler.Configuration);

                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_Success,
                        handler.Assembly,
                        handler.Type);
                }
            }
            catch (Exception ex)
            {
                string _message = string.Format(
                    CoreResources.Provisioning_Extensibility_Pipeline_Exception,
                    handler.Assembly,
                    handler.Type,
                    ex);
                Log.Error(_loggingSource, _message);
                throw new ExtensiblityPipelineException(_message, ex);

            }
        }

        /// <summary>
        /// Method to Invoke Custom Extraction Handlers. 
        /// </summary>
        /// <remarks>
        /// Ensure the ClientContext is not disposed in the custom provider.
        /// </remarks>
        /// <param name="ctx">Authenticated ClientContext that is passed to the custom provider.</param>
        /// <param name="handler">A custom Extensibility Provisioning Provider</param>
        /// <param name="template">ProvisioningTemplate that is passed to the custom provider</param>
        /// <param name="creationInformation">The Provisioning Template creation information object</param>
        /// <param name="scope">The PnPMonitoredScope of the current step in the pipeline</param>
        /// <exception cref="ExtensiblityPipelineException"></exception>
        /// <exception cref="ArgumentException">Provider.Assembly or Provider.Type is NullOrWhiteSpace></exception>
        /// <exception cref="ArgumentNullException">ClientContext is Null></exception>
        public ProvisioningTemplate ExecuteExtensibilityExtractionCallOut(ClientContext ctx, ExtensibilityHandler handler, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation, PnPMonitoredScope scope)
        {
            var _loggingSource = "OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensibilityManager.ExecuteCallout";

            if (ctx == null)
                throw new ArgumentNullException(CoreResources.Provisioning_Extensibility_Pipeline_ClientCtxNull);

            if (string.IsNullOrWhiteSpace(handler.Assembly))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);

            if (string.IsNullOrWhiteSpace(handler.Type))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_TypeName);

            ProvisioningTemplate parsedTemplate = null;

            try
            {

                var _instance = GetProviderInstance(handler);
                if (_instance is IProvisioningExtensibilityHandler)
                {
                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_BeforeInvocation,
                        handler.Assembly,
                        handler.Type);

                    parsedTemplate = (_instance as IProvisioningExtensibilityHandler).Extract(ctx, template, creationInformation, scope, handler.Configuration);

                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_Success,
                        handler.Assembly,
                        handler.Type);
                }
                else
                {
                    parsedTemplate = template;
                }
            }
            catch (Exception ex)
            {
                string _message = string.Format(
                    CoreResources.Provisioning_Extensibility_Pipeline_Exception,
                    handler.Assembly,
                    handler.Type,
                    ex);
                Log.Error(_loggingSource, _message);
                throw new ExtensiblityPipelineException(_message, ex);

            }

            return parsedTemplate;
        }
        private object GetProviderInstance(ExtensibilityHandler handler)
        {
            if (string.IsNullOrWhiteSpace(handler.Assembly))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);

            if (string.IsNullOrWhiteSpace(handler.Type))
                throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_TypeName);

            if (!handlerCache.ContainsKey(handler))
            {
#if NETSTANDARD2_0
                var _instance = Activator.CreateInstance(handler.GetType());
#else
                var _instance = Activator.CreateInstance(handler.Assembly, handler.Type).Unwrap();
#endif
                handlerCache.Add(handler, _instance);
            }
            return handlerCache[handler];
        }
    }
}
