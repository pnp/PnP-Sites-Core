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
    class ObjectExtensibilityProviders : ObjectHandlerBase
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
                foreach (var provider in template.Providers)
                {
                    if (provider.Enabled)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(provider.Configuration))
                            {
                                provider.Configuration = parser.ParseString(provider.Configuration);
                            }
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_Calling_tokenprovider_extensibility_callout__0_, provider.Assembly);
                            var _providedTokens = _extManager.ExecuteTokenProviderCallOut(context, provider, template);
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
                foreach (var provider in template.Providers)
                {
                    if (provider.Enabled)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(provider.Configuration))
                            {
                                //replace tokens in configuration data
                                provider.Configuration = parser.ParseString(provider.Configuration);
                            }
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ExtensibilityProviders_Calling_extensibility_callout__0_, provider.Assembly);
                            _extManager.ExecuteExtensibilityCallOut(context, provider, template);
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
                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {

            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Providers.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }
}
