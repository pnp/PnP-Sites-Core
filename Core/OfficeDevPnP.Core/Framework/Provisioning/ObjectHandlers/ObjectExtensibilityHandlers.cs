using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Extensibility Provider CallOut
    /// </summary>
    internal class ObjectExtensibilityHandlers : ObjectHandlerBase
    {
        ExtensibilityManager _extManager = new ExtensibilityManager();

        public override string Name
        {
            get { return "Extensibility Providers"; }

        }

        public TokenParser AddExtendedTokens(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                var handlers = applyingInformation != null ?
                    template.ExtensibilityHandlers.Union(applyingInformation.ExtensibilityHandlers) :
                    template.ExtensibilityHandlers;

                foreach (var handler in handlers)
                {
                    if (handler.Enabled)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(handler.Configuration))
                            {
                                handler.Configuration = parser.ParseString(handler.Configuration);
                            }
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_Calling_tokenprovider_extensibility_callout__0_, handler.Assembly);
                            var _providedTokens = _extManager.ExecuteTokenProviderCallOut(context, handler, template);
                            if (_providedTokens != null)
                            {
                                foreach (var token in _providedTokens)
                                {
                                    parser.AddToken(token);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_tokenprovider_callout_failed___0_____1_, ex.Message, ex.StackTrace);
                            throw;
                        }
                    }
                }
                return parser;
            }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;
                foreach (var handler in template.ExtensibilityHandlers
#pragma warning disable 618
                    .Union(template.Providers)
#pragma warning restore 618
                    .Union(applyingInformation.ExtensibilityHandlers))
                {
                    if (handler.Enabled)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(handler.Configuration))
                            {
                                //replace tokens in configuration data
                                handler.Configuration = parser.ParseString(handler.Configuration);
                            }
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_Calling_extensibility_callout__0_, handler.Assembly);
                            _extManager.ExecuteExtensibilityProvisionCallOut(context, handler, template, applyingInformation, parser, scope);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_callout_failed___0_____1_, ex.Message, ex.StackTrace);
                            throw;
                        }
                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;
                foreach (var handler in creationInfo.ExtensibilityHandlers)
                {
                    if (handler.Enabled)
                    {
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_Calling_extensibility_callout__0_, handler.Assembly);
                            template = _extManager.ExecuteExtensibilityExtractionCallOut(context, handler, template, creationInfo, scope);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_callout_failed___0_____1_, ex.Message, ex.StackTrace);
                            throw;
                        }
                    }
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {

            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
#pragma warning disable 618
                _willProvision = template.ExtensibilityHandlers
                    .Union(template.Providers)
                    .Union(applyingInformation.ExtensibilityHandlers).Any();
#pragma warning restore 618
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = creationInfo.ExtensibilityHandlers.Any();
            }
            return _willExtract.Value;
        }
    }
}
