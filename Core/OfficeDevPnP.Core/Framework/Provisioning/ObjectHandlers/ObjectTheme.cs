using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Feature = OfficeDevPnP.Core.Framework.Provisioning.Model.Feature;
using System;
using System.Linq;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities.Themes;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Utilities;
using Newtonsoft.Json;
using Microsoft.Online.SharePoint.TenantAdministration;
#if !ONPREMISES
using static OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities.TenantHelper;
#endif
using System.Runtime.Serialization;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectTheme : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Theme"; }
        }

        public override string InternalName => "Themes";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#if !ONPREMISES
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                var parsedName = parser.ParseString(template.Theme.Name);

                if (!string.IsNullOrEmpty(parsedName))
                {
                    web.EnsureProperty(w => w.Url);

                    if (Enum.TryParse<SharePointTheme>(parsedName, out SharePointTheme builtInTheme))
                    {
                        ThemeManager.ApplyTheme(web, builtInTheme);
                    }
                    else if (!string.IsNullOrEmpty(template.Theme.Palette))
                    {
                        var parsedPalette = parser.ParseString(template.Theme.Palette);

                        ThemeManager.ApplyTheme(web, parsedPalette, template.Theme.Name ?? parsedPalette);
                    }
                    else
                    {
                        //The account used for authenticating needs to be tenant administrator.
                        try
                        {
                            using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl()))
                            {
                                var tenant = new Tenant(tenantContext);
                                var theme = tenant.GetTenantTheme(parsedName);
                                tenantContext.Load(theme);
                                tenant.SetWebTheme(parsedName, web.Url);
                                tenantContext.ExecuteQueryRetry();
                            }
                        }
                        catch (Exception ex)
                        {
                            scope.LogWarning($"Custom theme could not be applied to site: {ex.Message}");
                            throw;
                        }
                    }
                }
            }
#endif
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return template.Theme != null;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return false;
        }
    }
}
