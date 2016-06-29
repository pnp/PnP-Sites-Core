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
using System.Threading.Tasks;
using System.Xml.Linq;

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
                    CleanupAllTestSiteCollections(tenantContext);
                }
            }
        }

        [TestInitialize()]
        public void Initialize()
        {
            sitecollectionName = sitecollectionNamePrefix + Guid.NewGuid().ToString();
        }

        #endregion

        #region Validation methods
        /// <summary>
        /// Validate two collection objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceElement"></param>
        /// <param name="targetElement"></param>
        /// <param name="props"></param>
        /// <returns></returns>
        public static bool ValidateObjects<T>(T sourceElement, T targetElement, List<string> property) where T : class
        {
            IEnumerable sElements = (IEnumerable)sourceElement;
            IEnumerable tElements = (IEnumerable)targetElement;
            int sCount = 0;
            int tCount = 0;

            foreach (string p in property)
            {
                foreach (object sElem in sElements)
                {
                    sCount++;
                    object sValue = sElem.GetType().GetProperty(p).GetValue(sElem);

                    foreach (object tElem in tElements)
                    {
                        object tValue = tElem.GetType().GetProperty(p).GetValue(tElem);

                        if (Convert.ToString(sValue) == Convert.ToString(tValue))
                        {
                            tCount++;
                            break;
                        }
                    }
                }
            }

            if (sCount != tCount)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceParser"></param>
        /// <param name="targetParser"></param>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        public static bool ValidateObjectSchemaXML<T>(TokenParser sourceParser, TokenParser targetParser, IEnumerable<T> source, IEnumerable<T> target, string property) where T : class
        {
            int scount = 0;
            int tcount = 0;

            foreach (var sField in source)
            {
                object sSchemaXml = sField.GetType().GetProperty("SchemaXml").GetValue(sField);
                XElement sourceElement = XElement.Parse(sourceParser.ParseString(sSchemaXml.ToString(), "~sitecollection", "~site"));
                var sValue = sourceElement.Attribute(property).Value;
                scount++;

                foreach (var tField in target)
                {
                    object tSchemaXml = sField.GetType().GetProperty("SchemaXml").GetValue(sField);
                    XElement targetElement = XElement.Parse(targetParser.ParseString(tSchemaXml.ToString(), "~sitecollection", "~site"));
                    var tValue = targetElement.Attribute(property).Value;

                    if (sValue == tValue)
                    {
                        tcount++;
                        break;
                    }
                }
            }

            if (scount != tcount)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="context"></param>
        /// <param name="security"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public static bool ValidateSecurity(ClientContext context, ObjectSecurity security, SecurableObject item)
        {
            int dataRowRoleAssignmentCount = security.RoleAssignments.Count;
            int roleCount = 0;

            IEnumerable roles = context.LoadQuery(item.RoleAssignments.Include(roleAsg => roleAsg.Member,
                roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name)));
            context.ExecuteQuery();

            foreach (var s in security.RoleAssignments)
            {
                foreach (Microsoft.SharePoint.Client.RoleAssignment r in roles)
                {
                    if (r.Member.LoginName.Contains(s.Principal) && r.RoleDefinitionBindings.Where(i => i.Name == s.RoleDefinition).FirstOrDefault() != null)
                    {
                        roleCount++;
                    }
                }
            }

            if (dataRowRoleAssignmentCount != roleCount)
            {
                return false;
            }

            return true;
        }
        #endregion

        #region Apply template and read the "result"
        public static Tuple<ProvisioningTemplate, ProvisioningTemplate> ApplyProvisioningTemplate(ClientContext cc, string templateName, ProvisioningTemplateApplyingInformation ptai=null, ProvisioningTemplateCreationInformation ptci = null)
        {
            // Read the template from XML and apply it
            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");
            ProvisioningTemplate sourceTemplate = provider.GetTemplate(templateName);

            if (ptai == null)
            {
                ptai = new ProvisioningTemplateApplyingInformation();
            }

            if (ptai.ProgressDelegate == null)
            {
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("Applying template - {0}/{1} - {2}", progress, total, message);
                };
            }

            sourceTemplate.Connector = provider.Connector;
            cc.Web.ApplyProvisioningTemplate(sourceTemplate, ptai);

            // Read the site we applied the template to 
            if (ptci == null)
            {
                ptci = new ProvisioningTemplateCreationInformation(cc.Web);
            }

            if (ptci.ProgressDelegate == null)
            {
                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("Getting template - {0}/{1} - {2}", progress, total, message);
                };
            }

            ProvisioningTemplate targetTemplate = cc.Web.GetProvisioningTemplate(ptci);

            return new Tuple<ProvisioningTemplate, ProvisioningTemplate>(sourceTemplate, targetTemplate);
        }
        #endregion

        #region Helper methods
#if !ONPREMISES
        internal static string CreateTestSiteCollection(Tenant tenant, string sitecollectionName)
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

        internal static string CreateTestSubSite(Tenant tenant, string sitecollectionUrl, string subSiteName)
        {
            // create a sub site in the central site collection
            Site site = tenant.GetSiteByUrl(sitecollectionUrl);
            tenant.Context.Load(site);
            tenant.Context.ExecuteQueryRetry();
            Web web = site.RootWeb;
            web.Context.Load(web);
            web.Context.ExecuteQueryRetry();

            //Create sub site
            SiteEntity sub = new SiteEntity() { Title = "Sub site for engine testing", Url = subSiteName, Description = "" };
            var subWeb = web.CreateWeb(sub);
            subWeb.EnsureProperty(t => t.Url);
            return subWeb.Url;
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

#endif

#if ONPREMISES
        private string CreateTestSiteCollection(Tenant tenant, string sitecollectionName)
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
