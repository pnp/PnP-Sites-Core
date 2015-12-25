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
    internal class ObjectWebSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Web Settings"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(
                    w => w.NoCrawl,
                    w => w.RequestAccessEmail,
                    w => w.MasterUrl,
                    w => w.CustomMasterUrl,
                    w => w.SiteLogoUrl,
                    w => w.RootFolder,
                    w => w.AlternateCssUrl);

                var webSettings = new WebSettings();
                webSettings.NoCrawl = web.NoCrawl;
                webSettings.RequestAccessEmail = web.RequestAccessEmail;
                webSettings.MasterPageUrl = web.MasterUrl;
                webSettings.CustomMasterPageUrl = web.CustomMasterUrl;
                webSettings.SiteLogo = web.SiteLogoUrl;
                webSettings.WelcomePage = web.RootFolder.WelcomePage;
                webSettings.AlternateCSS = web.AlternateCssUrl;
                template.WebSettings = webSettings;
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.WebSettings != null)
                {
                    var webSettings = template.WebSettings;
                    web.NoCrawl = webSettings.NoCrawl;
                    web.RequestAccessEmail = parser.ParseString(webSettings.RequestAccessEmail);
                    var masterUrl = parser.ParseString(webSettings.MasterPageUrl);
                    if (!string.IsNullOrEmpty(masterUrl))
                    {
                        web.MasterUrl = masterUrl;
                    }
                    var customMasterUrl = parser.ParseString(webSettings.CustomMasterPageUrl);
                    if (!string.IsNullOrEmpty(customMasterUrl))
                    {
                        web.CustomMasterUrl = customMasterUrl;
                    }
                    web.Description = parser.ParseString(webSettings.Description);
                    web.SiteLogoUrl = parser.ParseString(webSettings.SiteLogo);
                    var welcomePage = parser.ParseString(webSettings.WelcomePage);
                    if (!string.IsNullOrEmpty(welcomePage))
                    {
                        web.RootFolder.WelcomePage = welcomePage;
                        web.RootFolder.Update();
                    }
                    web.AlternateCssUrl = parser.ParseString(webSettings.AlternateCSS);

                    web.Update();
                    web.Context.ExecuteQueryRetry();
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return template.WebSettings != null;
        }
    }
}
