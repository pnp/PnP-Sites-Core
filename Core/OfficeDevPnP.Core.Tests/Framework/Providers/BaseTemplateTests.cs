using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Providers
{
    [TestClass]
    public class BaseTemplateTests
    {

        /// <summary>
        /// This is not a test, merely used to dump the needed template files
        /// </summary>
        [TestMethod]
        [Ignore]
        public void DumpBaseTemplates()
        {
            using (ClientContext ctx = TestCommon.CreateClientContext())
            {
                DumpTemplate(ctx, "STS#0");
                DumpTemplate(ctx, "BLOG#0");
                DumpTemplate(ctx, "BDR#0");
                DumpTemplate(ctx, "DEV#0");
                DumpTemplate(ctx, "OFFILE#1");

#if !CLIENTSDKV15
                DumpTemplate(ctx, "EHS#1");
                DumpTemplate(ctx, "BLANKINTERNETCONTAINER#0", "", "BLANKINTERNET#0");
#else
                DumpTemplate(ctx, "STS#1");
                DumpTemplate(ctx, "BLANKINTERNET#0");
#endif
                DumpTemplate(ctx, "BICENTERSITE#0");
                DumpTemplate(ctx, "SRCHCEN#0");
                DumpTemplate(ctx, "BLANKINTERNETCONTAINER#0", "CMSPUBLISHING#0");
                DumpTemplate(ctx, "ENTERWIKI#0");
                DumpTemplate(ctx, "PROJECTSITE#0");
                DumpTemplate(ctx, "COMMUNITY#0");
                DumpTemplate(ctx, "COMMUNITYPORTAL#0");
                DumpTemplate(ctx, "SRCHCENTERLITE#0");
                DumpTemplate(ctx, "VISPRUS#0");
            }
        }

        [TestMethod]
        [Ignore]
        public void DumpSingleTemplate()
        {
            using (ClientContext ctx = TestCommon.CreateClientContext())
            {
                DumpTemplate(ctx, "STS#0");
            }
        }

        private void DumpTemplate(ClientContext ctx, string template, string subSiteTemplate = "", string saveAsTemplate = "")
        {

            Uri devSiteUrl = new Uri(ConfigurationManager.AppSettings["SPODevSiteUrl"]);
            string baseUrl = String.Format("{0}://{1}", devSiteUrl.Scheme, devSiteUrl.DnsSafeHost);

            string siteUrl = "";
            if (subSiteTemplate.Length > 0)
            {
                siteUrl = string.Format("{1}/sites/template{0}/template{2}", template.Replace("#", ""), baseUrl, subSiteTemplate.Replace("#", ""));
#if !CLIENTSDKV15
                var siteCollectionUrl = string.Format("{1}/sites/template{0}", template.Replace("#", ""), baseUrl);
                CreateSiteCollection(template, siteCollectionUrl);
                using (var sitecolCtx = ctx.Clone(siteCollectionUrl))
                {
                    sitecolCtx.Web.Webs.Add(new WebCreationInformation()
                    {
                        Title = string.Format("template{0}", subSiteTemplate),
                        Language = 1033,
                        Url = string.Format("template{0}", subSiteTemplate.Replace("#", "")),
                        UseSamePermissionsAsParentSite = true
                    });
                    sitecolCtx.ExecuteQueryRetry();
                }
#endif
            }
            else
            {
                siteUrl = string.Format("{1}/sites/template{0}", template.Replace("#", ""), baseUrl);
#if !CLIENTSDKV15
                CreateSiteCollection(template, siteUrl);
#endif
            }

            using (ClientContext cc = ctx.Clone(siteUrl))
            {
                // Specify null as base template since we do want "everything" in this case
                ProvisioningTemplateCreationInformation creationInfo = new ProvisioningTemplateCreationInformation(cc.Web);
                creationInfo.BaseTemplate = null;

                // Override the save name. Case is online site collection provisioned using blankinternetcontainer#0 which returns
                // blankinternet#0 as web template using CSOM/SSOM API
                if (saveAsTemplate.Length > 0)
                {
                    template = saveAsTemplate;
                }

                ProvisioningTemplate p = cc.Web.GetProvisioningTemplate(creationInfo);
                if (subSiteTemplate.Length > 0)
                {
                    p.Id = String.Format("{0}template", subSiteTemplate.Replace("#", ""));
                }
                else
                {
                    p.Id = String.Format("{0}template", template.Replace("#", ""));
                }

                // Cleanup before saving
                p.Security.AdditionalAdministrators.Clear();


                XMLFileSystemTemplateProvider provider = new XMLFileSystemTemplateProvider(".", "");
                if (subSiteTemplate.Length > 0)
                {
                    provider.SaveAs(p, String.Format("{0}Template.xml", subSiteTemplate.Replace("#", "")));
                }
                else
                {
                    provider.SaveAs(p, String.Format("{0}Template.xml", template.Replace("#", "")));
                }

#if !CLIENTSDKV15
                using (var tenantCtx = TestCommon.CreateTenantClientContext())
                {
                    Tenant tenant = new Tenant(tenantCtx);
                    Console.WriteLine("Deleting new site {0}", string.Format("{1}/sites/template{0}", template.Replace("#", ""), baseUrl));
                    tenant.DeleteSiteCollection(siteUrl, false);
                }
#endif
            }
        }

#if !CLIENTSDKV15
        private static void CreateSiteCollection(string template, string siteUrl)
        {
            // check if site exists
            using (var tenantCtx = TestCommon.CreateTenantClientContext())
            {
                Tenant tenant = new Tenant(tenantCtx);

                if (tenant.SiteExists(siteUrl))
                {
                    Console.WriteLine("Deleting existing site {0}", siteUrl);
                    tenant.DeleteSiteCollection(siteUrl, false);
                }
                Console.WriteLine("Creating new site {0}", siteUrl);
                tenant.CreateSiteCollection(new Entities.SiteEntity()
                {
                    Lcid = 1033,
                    TimeZoneId = 4,
                    SiteOwnerLogin = (TestCommon.Credentials as SharePointOnlineCredentials).UserName,
                    Title = "Template Site",
                    Template = template,
                    Url = siteUrl,
                }, true, true);

            }
        }
#endif

        /// <summary>
        /// Get the base template for the current site
        /// </summary>
        [TestMethod]
        public void GetBaseTemplateForCurrentSiteTest()
        {
            using (ClientContext ctx = TestCommon.CreateClientContext())
            {
                ProvisioningTemplate t = ctx.Web.GetBaseTemplate();

                Assert.IsNotNull(t);
            }
        }

    }
}
