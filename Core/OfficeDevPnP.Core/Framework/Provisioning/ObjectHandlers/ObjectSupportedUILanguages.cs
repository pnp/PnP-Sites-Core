using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSupportedUILanguages : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Supported UI Languages"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {

                web.Context.Load(web, w => w.SupportedUILanguageIds);

                web.Context.ExecuteQueryRetry();

                foreach (var id in web.SupportedUILanguageIds)
                {
                    template.SupportedUILanguages.Add(new SupportedUILanguage() { LCID = id });
                }

            }

            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.IsMultilingual = true;
                web.Context.Load(web, w => w.SupportedUILanguageIds);
                web.Update();
                web.Context.ExecuteQueryRetry();

                var isDirty = false;

                foreach (var id in web.SupportedUILanguageIds)
                {
                    var found = template.SupportedUILanguages.Any(sl => sl.LCID == id);

                    if (!found)
                    {
                        web.RemoveSupportedUILanguage(id);
                        isDirty = true;
                    }
                }
                if (isDirty)
                {
                    web.Update();
                    web.Context.ExecuteQueryRetry();
                }

                foreach (var id in template.SupportedUILanguages)
                {
                    web.AddSupportedUILanguage(id.LCID);
                }
                web.Update();
                web.Context.Load(web, w => w.SupportedUILanguageIds);
                web.Context.ExecuteQueryRetry();
            }

            return parser;
        }
        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return template.SupportedUILanguages.Any();
        }
    }
}
