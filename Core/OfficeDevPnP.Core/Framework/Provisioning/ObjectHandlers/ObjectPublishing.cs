using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPublishing : ObjectHandlerBase
    {
        private const string AVAILABLEPAGELAYOUTS = "__PageLayouts";
        private const string DEFAULTPAGELAYOUT = "__DefaultPageLayout";
        private readonly Guid PUBLISHING_FEATURE = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
        public override string Name
        {
            get { return "Publishing"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope("Publishing"))
            {
                if (web.IsFeatureActive(PUBLISHING_FEATURE))
                {
                    var webTemplates = web.GetAvailableWebTemplates(web.Language, false);
                    web.Context.Load(webTemplates, wts => wts.Include(wt => wt.Name, wt => wt.Lcid));
                    web.Context.ExecuteQueryRetry();
                    Publishing publishing = new Publishing();
                    publishing.AvailableWebTemplates.AddRange(webTemplates.Select(wt => new AvailableWebTemplate() { TemplateName = wt.Name, LanguageCode = (int)wt.Lcid }));
                    publishing.AutoCheckRequirements = AutoCheckRequirementsOptions.MakeCompliant;

                    publishing.PageLayouts.AddRange(GetAvailablePageLayouts(web));

                    template.Publishing = publishing;
                }
            }
            return template;
        }


        private IEnumerable<PageLayout> GetAvailablePageLayouts(Web web)
        {
            var defaultLayoutXml = web.GetPropertyBagValueString(DEFAULTPAGELAYOUT, null);

            var defaultPageLayoutUrl = string.Empty;
            if (defaultLayoutXml != null)
            {
                defaultPageLayoutUrl = XElement.Parse(defaultLayoutXml).Attribute("url").Value;
            }

            List<PageLayout> layouts = new List<PageLayout>();

            var layoutsXml = web.GetPropertyBagValueString(AVAILABLEPAGELAYOUTS, null);

            var layoutsElement = XElement.Parse(layoutsXml);

            foreach (var layout in layoutsElement.Descendants("layout"))
            {
                if (layout.Attribute("url") != null)
                {
                    var pageLayout = new PageLayout();
                    pageLayout.Path = layout.Attribute("url").Value;
                    if (pageLayout.Path == defaultPageLayoutUrl)
                    {
                        pageLayout.IsDefault = true;
                    }
                    layouts.Add(pageLayout);
                }

            }
            return layouts;
        }




        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope("Publishing"))
            {

            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return web.IsFeatureActive(PUBLISHING_FEATURE);
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return template.Publishing != null;
        }
    }
}
