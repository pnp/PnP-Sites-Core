using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Configuration;
using OfficeDevPnP.Core.Entities;
using System.Threading;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Tests.AppModelExtensions
{
#if !ONPREMISES
    [TestClass()]
    public class TenantExtensionsTests
    {
        private static string sitecollectionNamePrefix = "TestPnPSC_123456789_";
        private string sitecollectionName = "";

        #region Test initialize and cleanup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                CleanupAllTestSiteCollections(tenantContext);
            }
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                CleanupAllTestSiteCollections(tenantContext);
            }
        }

        private static void CleanupAllTestSiteCollections(ClientContext tenantContext)
        {
            var tenant = new Tenant(tenantContext);

            try
            {
                var siteCols = tenant.GetSiteCollections();

                foreach (var siteCol in siteCols)
                {
                    if (siteCol.Url.Contains(sitecollectionNamePrefix))
                    {
                        try
                        {
                            // Drop the site collection from the recycle bin
                            if (tenant.CheckIfSiteExists(siteCol.Url, "Recycled"))
                            {
                                tenant.DeleteSiteCollectionFromRecycleBin(siteCol.Url, false);
                            }
                            else
                            {
                                // Eat the exceptions: would occur if the site collection is already in the recycle bin.
                                try
                                {
                                    // ensure the site collection in unlocked state before deleting
                                    tenant.SetSiteLockState(siteCol.Url, SiteLockState.Unlock);
                                }
                                catch { }

                                // delete the site collection, do not use the recyle bin
                                tenant.DeleteSiteCollection(siteCol.Url, false);
                            }
                        }
                        catch (Exception ex)
                        {
                            // eat all exceptions
                            Console.WriteLine(ex.ToDetailedString(tenant.Context));
                        }
                    }
                }
            }
            // catch exceptions with the GetSiteCollections call and log them so we can grab the corelation ID
            catch (Exception ex)
            {                
                Console.WriteLine(ex.ToDetailedString(tenant.Context));
                throw;
            }
        }
    

        [TestInitialize()]
        public void Initialize()
        {
            sitecollectionName = sitecollectionNamePrefix + Guid.NewGuid().ToString();
        }
        #endregion

        #region Get site collections tests
        [TestMethod()]
        [Timeout(15 * 60 * 1000)]
        public void GetSiteCollectionsTest()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);
                var siteCols = tenant.GetSiteCollections();

                Assert.IsTrue(siteCols.Any());

            }
        }

#if !NETSTANDARD2_0
        [TestMethod()]
        [Timeout(15 * 60 * 1000)]
        public void GetOneDriveSiteCollectionsTest()
        {
            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Web service tests are not supported when testing using app-only");
            }

            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);
                var siteCols = tenant.GetOneDriveSiteCollections();

                Assert.IsTrue(siteCols.Any());

            }
        }

        [TestMethod()]
        [Timeout(15 * 60 * 1000)]
        public void GetUserProfileServiceClientTest() {
            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Web service tests are not supported when testing using app-only");
            }

            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);
                var serviceClient = tenant.GetUserProfileServiceClient();
                tenantContext.Load(tenantContext.Web, w => w.CurrentUser);
                tenantContext.ExecuteQueryRetry();

                var profile = serviceClient.GetUserProfileByName(tenantContext.Web.CurrentUser.LoginName);

                Assert.IsNotNull(profile);
            }
        }
#endif
#endregion

        #region Site existance tests
        [TestMethod()]
        [Timeout(15 * 60 * 1000)]
        public void CheckIfSiteExistsTest() {
            using (var tenantContext = TestCommon.CreateTenantClientContext()) {
                var tenant = new Tenant(tenantContext);
                var siteCollections = tenant.GetSiteCollections();

                // Grab a random site collection from the list of returned site collections
                int siteNumberToCheck = new Random().Next(0, siteCollections.Count - 1);

                var site = siteCollections[siteNumberToCheck];
                var siteExists1 = tenant.CheckIfSiteExists(site.Url, "Active");
                Assert.IsTrue(siteExists1);

                try {
                    string devSiteUrl = TestCommon.AppSetting("SPODevSiteUrl");
                    string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, "aaabbbccc");
                    var siteExists2 = tenant.CheckIfSiteExists(siteToCreateUrl, "Active");
                    Assert.IsFalse(siteExists2, "Invalid site returned as valid.");
                }
                catch (ServerException) { }
            }
        }

        [TestMethod()]
        [Timeout(15 * 60 * 1000)]
        public void SiteExistsTest()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);
                var siteCollections = tenant.GetSiteCollections();

                var site = siteCollections.Last();
                var siteExists1 = tenant.SiteExists(site.Url);
                Assert.IsTrue(siteExists1);

                string devSiteUrl = TestCommon.AppSetting("SPODevSiteUrl");
                string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, "aaabbbccc");
                var siteExists2 = tenant.SiteExists(siteToCreateUrl);
                Assert.IsFalse(siteExists2, "Invalid site returned as valid.");
            }
        }

        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SubSiteExistsTest()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);
                string devSiteUrl = TestCommon.AppSetting("SPODevSiteUrl");
                Console.WriteLine("SubSiteExistsTest: step 1");
                string siteToCreateUrl = CreateTestSiteCollection(tenant, sitecollectionName);
                Console.WriteLine("SubSiteExistsTest: step 1.1");
                string subSiteUrlGood = "";
                string subSiteUrlWrong = "";

                Site site = tenant.GetSiteByUrl(siteToCreateUrl);
                tenant.Context.Load(site);
                tenant.Context.ExecuteQueryRetry();
                Web web = site.RootWeb;
                web.Context.Load(web);
                web.Context.ExecuteQueryRetry();
                Console.WriteLine("SubSiteExistsTest: step 1.2");

                //Create sub site
                SiteEntity sub = new SiteEntity() { Title = "Test Sub", Url = "sub", Description = "Test" };
                web.CreateWeb(sub);
                siteToCreateUrl = UrlUtility.EnsureTrailingSlash(siteToCreateUrl);
                subSiteUrlGood = String.Format("{0}{1}", siteToCreateUrl, sub.Url);
                subSiteUrlWrong = String.Format("{0}{1}", siteToCreateUrl, "8988980");

                // Check real sub site
                Console.WriteLine("SubSiteExistsTest: step 2");
                bool subSiteExists = tenant.SubSiteExists(subSiteUrlGood);
                Console.WriteLine("SubSiteExistsTest: step 2.1");
                Assert.IsTrue(subSiteExists);

                // check non existing sub site
                Console.WriteLine("SubSiteExistsTest: step 3");
                bool subSiteExists2 = tenant.SubSiteExists(subSiteUrlWrong);
                Console.WriteLine("SubSiteExistsTest: step 3.1");
                Assert.IsFalse(subSiteExists2);

                // check root site (= site collection). Will return true when existing
                Console.WriteLine("SubSiteExistsTest: step 4");
                bool subSiteExists3 = tenant.SubSiteExists(siteToCreateUrl);
                Console.WriteLine("SubSiteExistsTest: step 4.1");
                Assert.IsTrue(subSiteExists3);

                // check root site (= site collection) that does not exist. Will return false when non-existant
                Console.WriteLine("SubSiteExistsTest: step 5");
                bool subSiteExists4 = tenant.SubSiteExists(siteToCreateUrl + "8808809808");
                Console.WriteLine("SubSiteExistsTest: step 5.1");
                Assert.IsFalse(subSiteExists4);
            }
        }
        #endregion

        #region Site collection creation and deletion tests
        [TestMethod]
        [Timeout(45 * 60 * 1000)]
        public void CreateDeleteSiteCollectionTest()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);

                //Create site collection test
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 1");
                string siteToCreateUrl = CreateTestSiteCollection(tenant, sitecollectionName);
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 1.1");
                var siteExists = tenant.SiteExists(siteToCreateUrl);
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 1.2");
                Assert.IsTrue(siteExists, "Site collection creation failed");

                //Delete site collection test: move to recycle bin
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 2");
                tenant.DeleteSiteCollection(siteToCreateUrl, true);
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 2.1");
                bool recycled = tenant.CheckIfSiteExists(siteToCreateUrl, "Recycled");
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 2.2");
                Assert.IsTrue(recycled, "Site collection recycling failed");

                //Remove from recycle bin
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 3");
                tenant.DeleteSiteCollectionFromRecycleBin(siteToCreateUrl, true);
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 3.1");
                var siteExists2 = tenant.SiteExists(siteToCreateUrl);
                Console.WriteLine("CreateDeleteSiteCollectionTest: step 3.2");
                Assert.IsFalse(siteExists2, "Site collection deletion from recycle bin failed");
            }
        }

        [TestMethod]
        [Timeout(45 * 60 * 1000)]
        public void CreateDeleteCreateSiteCollectionTest()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);

                //Create site collection test
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 1");
                string siteToCreateUrl = CreateTestSiteCollection(tenant, sitecollectionName);
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 1.1");
                var siteExists = tenant.SiteExists(siteToCreateUrl);
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 1.1");
                Assert.IsTrue(siteExists, "Site collection creation failed");

                //Delete site collection test: move to recycle bin
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 2");
                tenant.DeleteSiteCollection(siteToCreateUrl, true);
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 2.1");
                bool recycled = tenant.CheckIfSiteExists(siteToCreateUrl, "Recycled");
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 2.2");
                Assert.IsTrue(recycled, "Site collection recycling failed");

                //Remove from recycle bin
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 3");
                tenant.DeleteSiteCollectionFromRecycleBin(siteToCreateUrl, true);
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 3.1");
                var siteExists2 = tenant.SiteExists(siteToCreateUrl);
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 3.2");
                Assert.IsFalse(siteExists2, "Site collection deletion from recycle bin failed");

                //Create a site collection using the same url as the previously deleted site collection
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 4");
                siteToCreateUrl = CreateTestSiteCollection(tenant, sitecollectionName);
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 4.1");
                siteExists = tenant.SiteExists(siteToCreateUrl);
                Console.WriteLine("CreateDeleteCreateSiteCollectionTest: step 4.2");
                Assert.IsTrue(siteExists, "Second site collection creation failed");
            }
        }
        #endregion

        #region Site lockstate tests
        [TestMethod]
        [Timeout(45 * 60 * 1000)]
        public void SetSiteLockStateTest()
        {
            try
            {
                using (var tenantContext = TestCommon.CreateTenantClientContext())
                {
                    tenantContext.RequestTimeout = 1000 * 60 * 15;

                    var tenant = new Tenant(tenantContext);
                    string devSiteUrl = TestCommon.AppSetting("SPODevSiteUrl");
                    string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);

                    Console.WriteLine("SetSiteLockStateTest: step 1");
                    if (!tenant.SiteExists(siteToCreateUrl))
                    {
                        siteToCreateUrl = CreateTestSiteCollection(tenant, sitecollectionName);
                        Console.WriteLine("SetSiteLockStateTest: step 1.1");
                        var siteExists = tenant.SiteExists(siteToCreateUrl);
                        Console.WriteLine("SetSiteLockStateTest: step 1.2");
                        Assert.IsTrue(siteExists, "Site collection creation failed");
                    }

                    Console.WriteLine("SetSiteLockStateTest: step 2");
                    // Set Lockstate NoAccess test
                    tenant.SetSiteLockState(siteToCreateUrl, SiteLockState.NoAccess, true);

                    Console.WriteLine("SetSiteLockStateTest: step 2.1");
                    var siteProperties = tenant.GetSitePropertiesByUrl(siteToCreateUrl, true);
                    Console.WriteLine("SetSiteLockStateTest: step 2.1");
                    tenantContext.Load(siteProperties);
                    tenantContext.ExecuteQueryRetry();
                    Assert.IsTrue(siteProperties.LockState == SiteLockState.NoAccess.ToString(), "LockState wasn't set to NoAccess");

                    // Set Lockstate NoAccess test
                    Console.WriteLine("SetSiteLockStateTest: step 3");
                    tenant.SetSiteLockState(siteToCreateUrl, SiteLockState.Unlock, true);
                    Console.WriteLine("SetSiteLockStateTest: step 3.1");
                    var siteProperties2 = tenant.GetSitePropertiesByUrl(siteToCreateUrl, true);
                    Console.WriteLine("SetSiteLockStateTest: step 3.2");
                    tenantContext.Load(siteProperties2);
                    tenantContext.ExecuteQueryRetry();
                    Assert.IsTrue(siteProperties2.LockState == SiteLockState.Unlock.ToString(), "LockState wasn't set to UnLock");

                    //Delete site collection, also
                    Console.WriteLine("SetSiteLockStateTest: step 4");
                    tenant.DeleteSiteCollection(siteToCreateUrl, false);
                    Console.WriteLine("SetSiteLockStateTest: step 4.1");
                    var siteExists2 = tenant.SiteExists(siteToCreateUrl);
                    Console.WriteLine("SetSiteLockStateTest: step 4.2");
                    Assert.IsFalse(siteExists2, "Site collection deletion, including from recycle bin, failed");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToDetailedString());
                throw;
            }
        }
        #endregion

        #region AppCatalog tests
        [TestMethod()]
        public void GetAppCatalogTest()
        {
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(tenantContext);
                Assert.IsNotNull(tenant.GetAppCatalog());
            }
        }
        #endregion

        #region Private helper methods
        private static string GetTestSiteCollectionName(string devSiteUrl, string siteCollection)
        {
            Uri u = new Uri(devSiteUrl);
            string host = String.Format("{0}://{1}", u.Scheme, u.DnsSafeHost);

            string path = u.AbsolutePath;
            if (path.EndsWith("/"))
            {
                path = path.Substring(0, path.Length - 1);
            }
            path = path.Substring(0, path.LastIndexOf('/'));

            return string.Format("{0}{1}/{2}", host, path, siteCollection);
        }

        private string CreateTestSiteCollection(Tenant tenant, string sitecollectionName)
        {
            try
            {
                string devSiteUrl = TestCommon.AppSetting("SPODevSiteUrl");
                string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);

                string siteOwnerLogin = TestCommon.AppSetting("SPOUserName");
                if (TestCommon.AppOnlyTesting())
                {
                    using (var clientContext = TestCommon.CreateClientContext())
                    {
                        List<UserEntity> admins = clientContext.Web.GetAdministrators();
                        siteOwnerLogin = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[2];
                    }
                }

                SiteEntity siteToCreate = new SiteEntity()
                {
                    Url = siteToCreateUrl,
                    Template = "STS#0",
                    Title = "Test",
                    Description = "Test site collection",
                    SiteOwnerLogin = siteOwnerLogin,
                };

                Console.WriteLine(String.Format("!!Before creating site collection {0}", siteToCreateUrl));
                tenant.CreateSiteCollection(siteToCreate, false, true);
                Console.WriteLine(String.Format("!!Site collection created {0}", siteToCreateUrl));
                return siteToCreateUrl;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToDetailedString(tenant.Context));
                throw;
            }
        }
        #endregion
    }
#endif
    }
