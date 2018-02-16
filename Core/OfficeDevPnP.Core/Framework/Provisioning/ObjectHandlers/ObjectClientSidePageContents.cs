using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
#if !ONPREMISES
    internal class ObjectClientSidePageContents: ObjectContentHandlerBase
    {
        public override string Name
        {
            get { return "Client Side Page Contents"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // This handler only extracts contents and adds them to the Files and Pages collection.
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Extract the Home Page
                web.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);

                var homePageUrl = web.RootFolder.WelcomePage;
                var homepageName = System.IO.Path.GetFileName(web.RootFolder.WelcomePage);
                if (string.IsNullOrEmpty(homepageName))
                {
                    homepageName = "Home.aspx";
                }

                // Extract the home page
                new ClientSidePageContentsHelper().ExtractClientSidePage(web, template, creationInfo, scope, homePageUrl, homepageName, true);

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

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }

#endif
}
