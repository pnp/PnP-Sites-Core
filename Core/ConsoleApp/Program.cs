using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "https://justdevelopment.sharepoint.com/blaat/subblaat/";
            string templateFile = "template_projectsite.xml";
            string path = @"D:\Projects\Caase.SiteProvisioning\Caase.SiteProvisioningWeb\Tenants\justdevelopment.sharepoint.com\Templates";

            AuthenticationManager authManager = new AuthenticationManager();
            using (ClientContext context = authManager.GetSharePointOnlineAuthenticatedContextTenant(
                    siteUrl,
                    "judith@justdevelopment.onmicrosoft.com",
                    "Steyne85j"))
            {
                Web web = context.Web;

                if (System.IO.File.Exists(path + "/" + templateFile))
                {

                    
                    XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(path, "");
                    ProvisioningTemplate template = provider.GetTemplate(templateFile);
                    template.Connector = provider.Connector;

                    ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();

                    // Taxonomies & Search zijn niet supported in AppOnly Authorization model
                    ptai.HandlersToProcess ^= OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.TermGroups;
                    ptai.HandlersToProcess ^= OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.SearchSettings;

                    web.ApplyProvisioningTemplate(template, ptai);
                }
            }
        }
    }
}
