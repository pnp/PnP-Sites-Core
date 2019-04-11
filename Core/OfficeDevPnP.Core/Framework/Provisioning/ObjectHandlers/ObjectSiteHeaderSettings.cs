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
#if !ONPREMISES
    internal class ObjectSiteHeaderSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Site Header"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(w => w.HeaderEmphasis, w => w.HeaderLayout, w => w.MegaMenuEnabled);
                var header = new SiteHeader();
                header.MenuStyle = web.MegaMenuEnabled ? SiteHeaderMenuStyle.MegaMenu : SiteHeaderMenuStyle.Cascading;
                switch (web.HeaderLayout)
                {
                    case HeaderLayoutType.Compact:
                        {
                            header.Layout = SiteHeaderLayout.Compact;
                            break;
                        }
                    case HeaderLayoutType.Standard:
                    default:
                        {
                            header.Layout = SiteHeaderLayout.Standard;
                            break;
                        }
                }

                if (Enum.TryParse<SiteHeaderBackgroundEmphasis>(web.HeaderEmphasis.ToString(), out SiteHeaderBackgroundEmphasis backgroundEmphasis))
                {
                    header.BackgroundEmphasis = backgroundEmphasis;
                }
                template.Header = header;
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Header != null)
                {
                    switch (template.Header.Layout)
                    {
                        case SiteHeaderLayout.Compact:
                            {
                                web.HeaderLayout = HeaderLayoutType.Compact;
                                break;
                            }
                        case SiteHeaderLayout.Standard:
                            {
                                web.HeaderLayout = HeaderLayoutType.Standard;
                                break;
                            }
                    }
                    web.HeaderEmphasis = (SPVariantThemeType)Enum.Parse(typeof(SPVariantThemeType), template.Header.BackgroundEmphasis.ToString());
                    web.MegaMenuEnabled = template.Header.MenuStyle == SiteHeaderMenuStyle.MegaMenu;
                    web.Update();
                    web.Context.ExecuteQueryRetry();
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            var baseTemplateValue = web.GetBaseTemplateId();
            if (baseTemplateValue.Equals("GROUP#0", StringComparison.InvariantCultureIgnoreCase) || baseTemplateValue.Equals("SITEPAGEPUBLISHING#0", StringComparison.InvariantCultureIgnoreCase) || baseTemplateValue.Equals("STS#3", StringComparison.InvariantCultureIgnoreCase))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            var baseTemplateValue = web.GetBaseTemplateId();
            if (baseTemplateValue.Equals("GROUP#0", StringComparison.InvariantCultureIgnoreCase) || baseTemplateValue.Equals("SITEPAGEPUBLISHING#0", StringComparison.InvariantCultureIgnoreCase) || baseTemplateValue.Equals("STS#3", StringComparison.InvariantCultureIgnoreCase))
            {
                return template.Header != null;
            }
            else
            {
                return false;
            }
        }
    }
#endif
}
