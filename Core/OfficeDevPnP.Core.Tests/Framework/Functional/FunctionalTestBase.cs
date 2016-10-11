using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass()]
    public abstract class FunctionalTestBase
    {
        private static string sitecollectionNamePrefix = "TestPnPSC_12345_";
        internal static string centralSiteCollectionUrl = "";
        internal static string centralSubSiteUrl = "";
        internal const string centralSubSiteName = "sub";
        internal static bool debugMode = false;

        private ProvisioningTemplate _sourceTemplate = null;
        private ProvisioningTemplate _targetTemplate = null;
        private TokenParser _sourceParser = null;
        private TokenParser _targetParser = null;
        internal string sitecollectionName = "";

        #region Test preparation
        public static void ClassInitBase(TestContext context)
        {
            // Drop all previously created site collections to keep the environment clean
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                if (!debugMode)
                {
                    CleanupAllTestSiteCollections(tenantContext);

                    // Each class inheriting from this base class gets a central test site collection, so let's create that one
                    var tenant = new Tenant(tenantContext);
                    centralSiteCollectionUrl = CreateTestSiteCollection(tenant, sitecollectionNamePrefix + Guid.NewGuid().ToString());

                    // Add a default sub site
                    centralSubSiteUrl = CreateTestSubSite(tenant, centralSiteCollectionUrl, centralSubSiteName);
                }
            }
        }

        public static void ClassCleanupBase()
        {
            if (!debugMode)
            {
                using (var tenantContext = TestCommon.CreateTenantClientContext())
                {
#if !ONPREMISES
                    CleanupAllTestSiteCollections(tenantContext);
#else
                    // first cleanup the just created one...most likely it's not indexed yet
                    try
                    {
                        Tenant t = new Tenant(tenantContext);
                        t.DeleteSiteCollection(centralSiteCollectionUrl);
                    }
                    catch { }

                    // Use search based site collection retreival to delete the one's that are left over from failed test cases
                    CleanupAllTestSiteCollections(tenantContext);
#endif
                }
            }
        }

        [TestInitialize()]
        public virtual void Initialize()
        {
            sitecollectionName = sitecollectionNamePrefix + Guid.NewGuid().ToString();
        }

        #endregion

        #region Apply template and read the "result"
        public TestProvisioningTemplateResult TestProvisioningTemplate(ClientContext cc, string templateName, Handlers handlersToProcess = Handlers.All, ProvisioningTemplateApplyingInformation ptai = null, ProvisioningTemplateCreationInformation ptci = null)
        {
            // Read the template from XML and apply it
            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");
            ProvisioningTemplate sourceTemplate = provider.GetTemplate(templateName);

            if (ptai == null)
            {
                ptai = new ProvisioningTemplateApplyingInformation();
                ptai.HandlersToProcess = handlersToProcess;
            }

            if (ptai.ProgressDelegate == null)
            {
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("Applying template - {0}/{1} - {2}", progress, total, message);
                };
            }

            sourceTemplate.Connector = provider.Connector;

            TokenParser sourceTokenParser = new TokenParser(cc.Web, sourceTemplate);

            cc.Web.ApplyProvisioningTemplate(sourceTemplate, ptai);

            // Read the site we applied the template to 
            if (ptci == null)
            {
                ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.HandlersToProcess = handlersToProcess;
            }

            if (ptci.ProgressDelegate == null)
            {
                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("Getting template - {0}/{1} - {2}", progress, total, message);
                };
            }

            ProvisioningTemplate targetTemplate = cc.Web.GetProvisioningTemplate(ptci);

            return new TestProvisioningTemplateResult()
            {
                SourceTemplate = sourceTemplate,
                SourceTokenParser = sourceTokenParser,
                TargetTemplate = targetTemplate,
                TargetTokenParser = new TokenParser(cc.Web, targetTemplate),
            };
        }
        #endregion

        #region Helper methods
#if !ONPREMISES
        internal static string CreateTestSiteCollection(Tenant tenant, string sitecollectionName)
        {
            try
            {
                string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];
                string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);

                string siteOwnerLogin = ConfigurationManager.AppSettings["SPOUserName"];
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

                tenant.CreateSiteCollection(siteToCreate, false, true);
                return siteToCreateUrl;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
        }

        private static void CleanupAllTestSiteCollections(ClientContext tenantContext)
        {
            var tenant = new Tenant(tenantContext);

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
                        Console.WriteLine(ex.ToString());
                    }
                }
            }
        }

        internal static string CreateTestSubSite(Tenant tenant, string sitecollectionUrl, string subSiteName)
        {
            // Create a sub site in the central site collection
            using (var cc = TestCommon.CreateClientContext(sitecollectionUrl))
            {
                //Create sub site
                SiteEntity sub = new SiteEntity() { Title = "Sub site for engine testing", Url = subSiteName, Description = "" };
                var subWeb = cc.Web.CreateWeb(sub);
                subWeb.EnsureProperty(t => t.Url);
                return subWeb.Url;
            }

            // Below approach is not working on edog...to be investigated
            //// create a sub site in the central site collection
            //Site site = tenant.GetSiteByUrl(sitecollectionUrl);
            //tenant.Context.Load(site);
            //tenant.Context.ExecuteQueryRetry();
            //Web web = site.RootWeb;
            //web.Context.Load(web);
            //web.Context.ExecuteQueryRetry();

            ////Create sub site
            //SiteEntity sub = new SiteEntity() { Title = "Sub site for engine testing", Url = subSiteName, Description = "" };
            //var subWeb = web.CreateWeb(sub);
            //subWeb.EnsureProperty(t => t.Url);
            //return subWeb.Url;
        }
#else
        private static string CreateTestSiteCollection(Tenant tenant, string sitecollectionName)
        {
            string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];

            string siteOwnerLogin = string.Format("{0}\\{1}", ConfigurationManager.AppSettings["OnPremDomain"], ConfigurationManager.AppSettings["OnPremUserName"]);
            if (TestCommon.AppOnlyTesting())
            {
                using (var clientContext = TestCommon.CreateClientContext())
                {
                    List<UserEntity> admins = clientContext.Web.GetAdministrators();
                    siteOwnerLogin = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[1];
                }
            }

            string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);
            SiteEntity siteToCreate = new SiteEntity()
            {
                Url = siteToCreateUrl,
                Template = "STS#0",
                Title = "Test",
                Description = "Test site collection",
                SiteOwnerLogin = siteOwnerLogin,
            };

            tenant.CreateSiteCollection(siteToCreate);
            return siteToCreateUrl;
        }

        internal static string CreateTestSubSite(Tenant tenant, string sitecollectionUrl, string subSiteName)
        {
            using (ClientContext cc = new ClientContext(sitecollectionUrl))
            {
                //Create sub site
                SiteEntity sub = new SiteEntity() { Title = "Sub site for engine testing", Url = subSiteName, Description = "" };
                var subWeb = cc.Web.CreateWeb(sub);
                subWeb.EnsureProperty(t => t.Url);
                return subWeb.Url;
            }
        }

        private static void CleanupAllTestSiteCollections(ClientContext tenantContext)
        {
            string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];
                       
            var tenant = new Tenant(tenantContext);
            try
            {
                using (ClientContext cc = tenantContext.Clone(devSiteUrl))
                {
                    var sites = cc.Web.SiteSearch();

                    foreach(var site in sites)
                    {
                        if (site.Url.ToLower().Contains(sitecollectionNamePrefix.ToLower()))
                        {
                            tenant.DeleteSiteCollection(site.Url);
                        }
                    }
                }
            }
            catch
            { }
        }

        private void CleanupCreatedTestSiteCollections(ClientContext tenantContext)
        {
            string devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];
            String testSiteCollection = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);

            //Ensure the test site collection was deleted and removed from recyclebin
            var tenant = new Tenant(tenantContext);
            try
            {
                tenant.DeleteSiteCollection(testSiteCollection);
            }
            catch
            { }
        }
#endif

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
        #endregion

    }
}
