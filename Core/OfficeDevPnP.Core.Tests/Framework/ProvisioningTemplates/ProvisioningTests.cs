#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using Configuration = OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.ProvisioningTemplates
{
    [TestClass]
    public class ProvisioningTests
    {

        [TestMethod]
        public void GetGroupInfoTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                OfficeDevPnP.Core.Sites.SiteCollection.GetGroupInfo(context, "demo1").GetAwaiter().GetResult();
            }
        }

        //[TestMethod]
        //public void GetTenantTemplateTest()
        //{
        //    using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, string.Join(" ", scope)))))
        //    {
        //        using (var context = TestCommon.CreateTenantClientContext())
        //        {
        //            var tenant = new Tenant(context);
        //            var configuration = new ExtractConfiguration();
        //            //configuration.Tenant.Sequence = new Configuration.Tenant.Sequence.ExtractSequenceConfiguration()
        //            //{
        //            //    IncludeJoinedSites = true,
        //            //    IncludeSubsites = true,
        //            //    MaxSubsiteDepth = 2,
        //            //    SiteUrls = { "https://erwinmcm.sharepoint.com/sites/demo1" }
        //            //};

        //            configuration.Tenant.Teams = new Configuration.Tenant.Teams.ExtractTeamsConfiguration()
        //            {
        //                TeamSiteUrls = { "https://erwinmcm.sharepoint.com/sites/teamchild" },
        //                IncludeMessages = true
        //            };
                    
        //            //      configuration.Handlers.Add(ConfigurationHandler.Lists);
        //            //configuration.Handlers.Add(ConfigurationHandler.WebSettings);

        //            //configuration.Lists.Lists.Add(new Core.Framework.Provisioning.Model.Configuration.Lists.Lists.ExtractConfiguration()
        //            //{
        //            //    Title = "Test"
        //            //});
        //            //configuration.ProgressAction = (message, step, total) =>
        //            //{
        //            //    Trace.Write($"{step}|{total}|{message}");
        //            //};
        //            var tenantTemplate = new SiteToTemplateConversion().GetTenantTemplate(tenant, configuration);
        //        }
        //    }
        //}

        [TestMethod]
        public void ProvisionTenantTemplate()
        {
            var resourceFolder = string.Format(@"{0}\..\..\Resources\Templates", AppDomain.CurrentDomain.BaseDirectory);
            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(resourceFolder, "");

            var existingTemplate = provider.GetTemplate("ProvisioningSchema-2018-07-FullSample-01.xml");

            Guid siteGuid = Guid.NewGuid();
            int siteId = siteGuid.GetHashCode();
            var template = new ProvisioningTemplate();
            template.Id = "TestTemplate";
            template.Lists.Add(new ListInstance()
            {
                Title = "Testlist",
                TemplateType = 100,
                Url = "lists/testlist"
            });

            template.TermGroups.AddRange(existingTemplate.TermGroups);

            ProvisioningHierarchy hierarchy = new ProvisioningHierarchy();

            hierarchy.Templates.Add(template);

            hierarchy.Parameters.Add("CompanyName", "Contoso");

            var sequence = new ProvisioningSequence();

            sequence.TermStore = new ProvisioningTermStore();
            var termGroup = new TermGroup() { Name = "Contoso TermGroup" };
            var termSet = new TermSet() { Name = "Projects", Id = Guid.NewGuid(), IsAvailableForTagging = true, Language = 1033 };
            var term = new Term() { Name = "Contoso Term" };

            termSet.Terms.Add(term);
            // termGroup.TermSets.Add(termSet);

            var existingTermSet = existingTemplate.TermGroups[0].TermSets[0];
            termGroup.TermSets.Add(existingTermSet);

            // sequence.TermStore.TermGroups.Add(termGroup);

            var teamSite1 = new TeamSiteCollection()
            {
                //  Alias = $"prov-1-{siteId}",
                Alias = "prov-1",
                Description = "prov-1",
                DisplayName = "prov-1",
                IsHubSite = false,
                IsPublic = false,
                Title = "prov-1",
            };
            teamSite1.Templates.Add("TestTemplate");

            var subsite = new TeamNoGroupSubSite()
            {
                Description = "Test Sub",
                Url = "testsub1",
                Language = 1033,
                TimeZoneId = 4,
                Title = "Test Sub",
                UseSamePermissionsAsParentSite = true
            };
            subsite.Templates.Add("TestTemplate");
            teamSite1.Sites.Add(subsite);

            sequence.SiteCollections.Add(teamSite1);

            var teamSite2 = new TeamSiteCollection()
            {
                Alias = $"prov-2-{siteId}",
                Description = "prov-2",
                DisplayName = "prov-2",
                IsHubSite = false,
                IsPublic = false,
                Title = "prov-2"
            };
            teamSite2.Templates.Add("TestTemplate");

            sequence.SiteCollections.Add(teamSite2);

            hierarchy.Sequences.Add(sequence);


            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var applyingInformation = new ProvisioningTemplateApplyingInformation();
                applyingInformation.ProgressDelegate = (message, step, total) =>
                {
                    if (message != null)
                    {


                    }
                };

                var tenant = new Tenant(tenantContext);

                tenant.ApplyProvisionHierarchy(hierarchy, sequence.ID, applyingInformation);
            }
        }
    }
}
#endif