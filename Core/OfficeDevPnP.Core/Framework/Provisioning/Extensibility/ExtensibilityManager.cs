using System;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Reflection;
using System.IO;
using System.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    ///  Provisioning Framework Component that is used for invoking custom providers during the provisioning process.
    /// </summary>
    public partial class ExtensibilityManager
    {
        private Dictionary<ExtensibilityHandler, Object> handlerCache = new Dictionary<ExtensibilityHandler, Object>();
		private static List<Tuple<Assembly, byte[], byte[]>> packedAssemblyCache = null;

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
		public IEnumerable<TokenDefinition> ExecuteTokenProviderCallOut(ClientContext ctx, ExtensibilityHandler provider, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation = null)
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

				var _instance = GetProviderInstance(provider, template.Connector, 
					applyingInformation != null ? applyingInformation.CustomAssemblyBinding : CustomAssemblyBindingPolicy.None) as IProvisioningExtensibilityTokenProvider;
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

				var _instance = GetProviderInstance(handler, template.Connector, applyingInformation.CustomAssemblyBinding);
				if (_instance is IProvisioningExtensibilityProvider)
                {
                    Log.Info(_loggingSource,
                        CoreResources.Provisioning_Extensibility_Pipeline_BeforeInvocation,
                        handler.Assembly,
                        handler.Type);

                    (_instance as IProvisioningExtensibilityProvider).ProcessRequest(ctx, template, handler.Configuration);

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

		private object GetProviderInstance(ExtensibilityHandler handler, FileConnectorBase provider = null, CustomAssemblyBindingPolicy bindingPocily = CustomAssemblyBindingPolicy.None)
		{
			if (string.IsNullOrWhiteSpace(handler.Assembly))
				throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);

			if (string.IsNullOrWhiteSpace(handler.Type))
				throw new ArgumentException(CoreResources.Provisioning_Extensibility_Pipeline_Missing_TypeName);

			if (!handlerCache.ContainsKey(handler))
			{
				var typeName = Assembly.CreateQualifiedName(handler.Assembly, handler.Type);
				var type = GetType(typeName, provider, bindingPocily);
				var _instance = type != null ? Activator.CreateInstance(type) : null;
				handlerCache.Add(handler, _instance);
			}
			return handlerCache[handler];
		}

		internal static Type GetType(string handlerType, FileConnectorBase connector, CustomAssemblyBindingPolicy bindingPolicy = CustomAssemblyBindingPolicy.None)
		{
			Type res = Type.GetType(handlerType, false);
			if (res == null && bindingPolicy != CustomAssemblyBindingPolicy.None)
			{
				if (packedAssemblyCache == null)
				{
					packedAssemblyCache = connector != null ? GetAssembliesFromConnector(connector, bindingPolicy) :
						new List<Tuple<Assembly, byte[], byte[]>>();

					if (packedAssemblyCache?.Count > 0)
					{
						AppDomain.CurrentDomain.AssemblyResolve += ResolveProviderAssembly;
					}
				}
				res = Type.GetType(handlerType, false);
			}
			return res;
		}

		private static List<Tuple<Assembly, byte[], byte[]>> GetAssembliesFromConnector(FileConnectorBase connector, CustomAssemblyBindingPolicy policy)
		{
			List<Tuple<Assembly, byte[], byte[]>> res = new List<Tuple<Assembly, byte[], byte[]>>();
			var files = connector.GetFiles();
			foreach (var file in files)
			{
				var ext = Path.GetExtension(file);
				if (string.Compare(ext, ".dll", true) == 0)
				{
					var fileName = Path.GetFileName(file);
					var assemblyStream = connector.GetFileStream(fileName);
					byte[] assemblyBytes = null;
					using (BinaryReader br = new BinaryReader(assemblyStream))
					{
						assemblyBytes = br.ReadBytes((int)assemblyStream.Length);
					}

					var assembly = Assembly.ReflectionOnlyLoad(assemblyBytes);
					var hasStrongName = assembly.Evidence.OfType<System.Security.Policy.StrongName>().Any();
					var signed = assembly.Evidence.OfType<System.Security.Policy.Publisher>().Any();

					if ((!signed && ((policy & CustomAssemblyBindingPolicy.Signed) != 0)) ||
						(!hasStrongName && ((policy & CustomAssemblyBindingPolicy.StrongName) != 0)))
					{
						continue;
					}

					var symbolsStream = connector.GetFileStream(assembly.GetName().Name + ".pdb");
					byte[] symbolsBytes = null;
					if (symbolsStream != null)
					{
						using (BinaryReader br = new BinaryReader(symbolsStream))
						{
							symbolsBytes = br.ReadBytes((int)symbolsStream.Length);
						}
					}
					res.Add(new Tuple<Assembly, byte[], byte[]>(assembly, assemblyBytes, symbolsBytes));
				}
			}
			return res;
		}

		internal static Tuple<string, AssemblyName> ParseTypeName(string name)
		{
			Tuple<string, AssemblyName> res = null;

			//check if type is already loaded
			var type = Type.GetType(name, false);
			if (type != null)
			{
				res = new Tuple<string, AssemblyName>(type.FullName, type.Assembly.GetName());
			}
			else
			{
				var _loggingSource = "OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensibilityManager.ParseTypeName";
				//type maybe loaded later if custom assembly binding enabled
				try
				{
					var index = name.IndexOf(',');
					if (index > 0)
					{
						res = new Tuple<string, AssemblyName>(name.Substring(0, index).Trim(), new AssemblyName(name.Substring(index + 1).Trim()));
					}
				}
				catch (Exception e)
				{
					Log.Error(e, _loggingSource, CoreResources.Provisioning_Extensibility_Pipeline_Missing_AssemblyName);
				}
			}
			return res;
		}

		private static Assembly ResolveProviderAssembly(object sender, ResolveEventArgs e)
		{
			Assembly res = null;
			var domain = AppDomain.CurrentDomain;
			var assemblies = domain.GetAssemblies();
			var name = new AssemblyName(e.Name);
			var hasFullName = name.CultureInfo != null && name.Version != null;

			//try resolve from loaded assemblies
			res = assemblies.FirstOrDefault(a => hasFullName ?
				string.Equals(name.FullName, a.FullName, StringComparison.InvariantCultureIgnoreCase) :
				string.Equals(name.Name, a.GetName().Name, StringComparison.InvariantCultureIgnoreCase));

			if (res == null && packedAssemblyCache != null && packedAssemblyCache.Count > 0)
			{
				//try resolve from cached assemblies
				var assembly = packedAssemblyCache.FirstOrDefault(a => hasFullName ?
				string.Equals(name.FullName, a.Item1.FullName, StringComparison.InvariantCultureIgnoreCase) :
				string.Equals(name.Name, a.Item1.GetName().Name, StringComparison.InvariantCultureIgnoreCase));
				if (assembly != null)
				{
					res = Assembly.Load(assembly.Item2, assembly.Item3);
				}
			}
			return res;
		}
	}
}
