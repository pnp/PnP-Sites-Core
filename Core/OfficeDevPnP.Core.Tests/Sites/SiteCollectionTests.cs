using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Sites
{
#if !ONPREMISES
    [TestClass]
    public class SiteCollectionTests
    {
        private string communicationSiteGuid;
        private string teamSiteGuid;
        private string baseUrl;

        [TestInitialize]
        public void Initialize()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                communicationSiteGuid = Guid.NewGuid().ToString("N");
                teamSiteGuid = Guid.NewGuid().ToString("N");
                var baseUri = new Uri(clientContext.Url);
                baseUrl = $"{baseUri.Scheme}://{baseUri.Host}:{baseUri.Port}";
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var clientContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(clientContext);
                tenant.DeleteSiteCollection($"{baseUrl}/sites/site{communicationSiteGuid}", false);
                // Commented this, first group cleanup needs to be implemented in this test case
                //tenant.DeleteSiteCollection($"{baseUrl}/sites/site{teamSiteGuid}", false);
                //TODO: Cleanup group
            }
        }

        [TestMethod]
        public async Task CreateCommunicationSiteTestAsync()
        {

            using (var clientContext = TestCommon.CreateClientContext())
            {

                var commResults = await clientContext.CreateSiteAsync(new Core.Sites.CommunicationSiteCollectionCreationInformation()
                {
                    Url = $"{baseUrl}/sites/site{communicationSiteGuid}",
                    SiteDesign = Core.Sites.CommunicationSiteDesign.Blank,
                    Title = "Comm Site Test",
                    Lcid = 1033
                });

                Assert.IsNotNull(commResults);
            }
        }

        //[TestMethod]
        //public async Task CreateTeamSiteTestAsync()
        //{
        //    using (var clientContext = TestCommon.CreateClientContext())
        //    {
        //        var teamResults = await clientContext.CreateSiteAsync(new Core.Sites.TeamSiteCollectionCreationInformation()
        //        {
        //            Alias = $"site{teamSiteGuid}",
        //            DisplayName = "Team Site Test",
        //        });
        //        Assert.IsNotNull(teamResults);
        //    }

        //}
    }
#endif
}
