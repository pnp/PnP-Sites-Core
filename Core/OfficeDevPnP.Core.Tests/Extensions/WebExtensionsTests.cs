using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Tests;
using System.IO;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Entities;

namespace Microsoft.SharePoint.Client.Tests
{
    [TestClass()]
    public class WebExtensionsTests
    {
        const string INDEXED_PROPERTY_KEY = "vti_indexedpropertykeys";
        private string _key = null;
        private string _value_string = null;
        private int _value_int = 12345;
        const string APPNAME = "HelloWorldApp";
        const string contentTypeName = "PnP Test Content Type";
        const string contentTypeGroupName = "PnP Web Extensions Test";
        private ClientContext clientContext;

        #region Test initialize and cleanup
        [TestInitialize()]
        public void Initialize()
        {
            clientContext = TestCommon.CreateClientContext();

            _key = "TEST_KEY_" + DateTime.Now.ToFileTime();
            _value_string = "TEST_VALUE_" + DateTime.Now.ToFileTime();

            // Activate sideloading in order to test apps
            clientContext.Load(clientContext.Site, s => s.Id);
            clientContext.ExecuteQueryRetry();
            clientContext.Site.ActivateFeature(Constants.FeatureId_Site_AppSideLoading);

            var provisionTemplate = new ProvisioningTemplate();
            var contentType = new OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType()
            {
                Id = "0x010100503B9E20E5455344BFAC2292DC6FAB81",
                Name = contentTypeName,
                Group = contentTypeGroupName,
                Description = "Test Description",
                Overwrite = true,
                Hidden = false
            };

            provisionTemplate.ContentTypes.Add(contentType);
            TokenParser parser = new TokenParser(clientContext.Web, provisionTemplate);
            new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(clientContext.Web, provisionTemplate, parser,
                new ProvisioningTemplateApplyingInformation());
        }

        [TestCleanup()]
        public void Cleanup()
        {
            // Deactivate sideloading
            clientContext.Load(clientContext.Site);
            clientContext.ExecuteQueryRetry();
            clientContext.Site.DeactivateFeature(Constants.FeatureId_Site_AppSideLoading);

            var props = clientContext.Web.AllProperties;
            clientContext.Load(props);
            clientContext.ExecuteQueryRetry();

            // Implement cleanup mechanism that cleans test stranglers + also cleans up the NoMobileMapping key that's generated per created sub site
            List<string> keysToDelete = new List<string>(10);
            foreach (var prop in props.FieldValues)
            {
                if (prop.Key.StartsWith("TEST_KEY_", StringComparison.InvariantCultureIgnoreCase) ||
                    prop.Key.StartsWith("TEST_VALUE_", StringComparison.InvariantCultureIgnoreCase) ||
                    prop.Key.StartsWith("__NoMobileMapping", StringComparison.InvariantCultureIgnoreCase))
                {
                    keysToDelete.Add(prop.Key);
                }
            }

            int batch = 0;
            foreach (string key in keysToDelete)
            {
                props[key] = null;
                props.FieldValues.Remove(key);
                batch++;

                // send cleanup in batches of 50 to the server
                if (batch >= 50)
                {
                    clientContext.Web.Update();
                    clientContext.ExecuteQueryRetry();
                    batch = 0;
                }
            }
            clientContext.Web.Update();
            clientContext.ExecuteQueryRetry();

            var instances = AppCatalog.GetAppInstances(clientContext, clientContext.Web);
            clientContext.Load(instances);
            clientContext.ExecuteQueryRetry();

            string appToRemove = APPNAME;
#if ONPREMISES
            appToRemove += "15";
#endif

            foreach (var instance in instances)
            {
                if (string.Equals(instance.Title, appToRemove, StringComparison.OrdinalIgnoreCase))
                {
                    instance.Uninstall();
                    clientContext.ExecuteQueryRetry();
                    break;
                }
            }

            var ct = clientContext.Web.GetContentTypeByName(contentTypeName);
            if (ct != null)
            {
                ct.DeleteObject();
                clientContext.ExecuteQueryRetry();
            }

            bool dirty = false;
            clientContext.Load(clientContext.Web.Webs, wc => wc.Include(w => w.Title, w => w.ServerRelativeUrl));
            clientContext.ExecuteQueryRetry();
            foreach (Web subWeb in clientContext.Web.Webs.ToList())
            {
                if (subWeb.Title.StartsWith("Test_") || subWeb.ServerRelativeUrl.Contains("Test_"))
                {
                    subWeb.DeleteObject();
                    dirty = true;
                }
            }
            if (dirty)
            {
                clientContext.ExecuteQueryRetry();
            }

            clientContext.Dispose();
        }
        #endregion

        #region Property bag tests
        [TestMethod()]
        public void SetPropertyBagValueIntTest()
        {
            clientContext.Web.SetPropertyBagValue(_key, _value_int);

            var props = clientContext.Web.AllProperties;
            clientContext.Load(props);
            clientContext.ExecuteQueryRetry();
            Assert.IsTrue(props.FieldValues.ContainsKey(_key));
            Assert.AreEqual(_value_int, props.FieldValues[_key] as int?);
        }

        [TestMethod()]
        public void SetPropertyBagValueStringTest()
        {
            clientContext.Web.SetPropertyBagValue(_key, _value_string);

            var props = clientContext.Web.AllProperties;
            clientContext.Load(props);
            clientContext.ExecuteQueryRetry();
            Assert.IsTrue(props.FieldValues.ContainsKey(_key), "Entry not added");
            Assert.AreEqual(_value_string, props.FieldValues[_key] as string, "Entry not set with correct value");
        }

        [TestMethod()]
        public void SetPropertyBagValueHandlesLocalPropertyCacheTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            string noUpdateKey = _key + "_NoUpdate_" + DateTime.Now.ToFileTime();
            props[noUpdateKey] = "This key is never added to the server because no web.Update() is called. This leads to future client caching issues if Web.AllProperties.RefreshLoad() or web.Context.Load(web.AllProperties) are called.";
            web.Context.ExecuteQueryRetry();

            clientContext.Web.SetPropertyBagValue(_key, _value_string);

            props.FieldValues.Remove(noUpdateKey); // Need to remove this key before refreshing from server.

            clientContext.Load(props);
            clientContext.ExecuteQueryRetry();

            Assert.IsTrue(props.FieldValues.ContainsKey(_key), "Entry not added");
            Assert.AreEqual(_value_string, props.FieldValues[_key] as string, "Entry not set with correct value");
            Assert.IsFalse(props.FieldValues.ContainsKey(noUpdateKey), "The key '" + noUpdateKey + "' should not exist on the server");
        }

        [TestMethod()]
        public void SetPropertyBagValueMultipleRunsTest()
        {
            string key2 = _key + "_multiple";
            clientContext.Web.SetPropertyBagValue(key2, _value_string);
            clientContext.Web.SetPropertyBagValue(_key, _value_string);

            var props = clientContext.Web.AllProperties;
            clientContext.Load(props);
            clientContext.ExecuteQueryRetry();
            Assert.IsTrue(props.FieldValues.ContainsKey(_key), "Entry not added");
            Assert.AreEqual(_value_string, props.FieldValues[_key] as string, "Entry not set with correct value");
        }

        [TestMethod()]
        public void RemovePropertyBagValueTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();

            props[_key] = _value_string;

            web.Update();
            web.Context.ExecuteQueryRetry();

            web.RemovePropertyBagValue(_key);

            props.RefreshLoad();
            clientContext.ExecuteQueryRetry();
            Assert.IsFalse(props.FieldValues.ContainsKey(_key), "Entry not removed");
        }

        [TestMethod()]
        public void GetPropertyBagValueIntTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();

            props[_key] = _value_int;

            web.Update();
            web.Context.ExecuteQueryRetry();

            var intValue = web.GetPropertyBagValueInt(_key, -1);

            Assert.IsInstanceOfType(intValue, typeof(int?), "No int value returned");
            Assert.AreEqual(_value_int, intValue, "Incorrect value returned");

            // Check for non-existing key
            intValue = web.GetPropertyBagValueInt("_key_" + DateTime.Now.ToFileTime(), -12345);
            Assert.IsInstanceOfType(intValue, typeof(int?), "No int value returned");
            Assert.AreEqual(-12345, intValue, "Incorrect value returned");
        }

        [TestMethod()]
        public void GetPropertyBagValueStringTest()
        {
            var notExistingKey = "NOTEXISTINGKEY";
            var web = clientContext.Web;
            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();

            props[_key] = _value_string;

            web.Update();
            web.Context.ExecuteQueryRetry();

            var stringValue = web.GetPropertyBagValueString(_key, notExistingKey);

            Assert.IsInstanceOfType(stringValue, typeof(string), "No string value returned");
            Assert.AreEqual(_value_string, stringValue, "Incorrect value returned");

            // Check for non-existing key
            stringValue = web.GetPropertyBagValueString("_key_" + DateTime.Now.ToFileTime(), notExistingKey);
            Assert.IsInstanceOfType(stringValue, typeof(string), "No string value returned");
            Assert.AreEqual(notExistingKey, stringValue, "Incorrect value returned");
        }

        [TestMethod()]
        public void GetPropertyBagValueHandlesLocalPropertyCacheTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            string noUpdateKey = _key + "_NoUpdate_" + DateTime.Now.ToFileTime();

            props[_key] = _value_string;
            web.Update();

            props[noUpdateKey] = "This key is never added to the server because no web.Update() is called. This leads to future client caching issues if Web.AllProperties.RefreshLoad() or web.Context.Load(web.AllProperties) are called.";
            web.Context.ExecuteQueryRetry();

            var stringValue = web.GetPropertyBagValueString(_key, _value_string);

            Assert.IsInstanceOfType(stringValue, typeof(string), "No string value returned");
            Assert.AreEqual(_value_string, stringValue, "Incorrect value returned");

            props.FieldValues.Remove(noUpdateKey); // Need to remove this key before refreshing from server

            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();

            Assert.IsFalse(props.FieldValues.ContainsKey(noUpdateKey), "The key '" + noUpdateKey + "' should not exist on the server");
        }

        [TestMethod()]
        public void PropertyBagContainsKeyTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();

            props[_key] = _value_string;

            web.Update();
            web.Context.ExecuteQueryRetry();

            Assert.IsTrue(web.PropertyBagContainsKey(_key));
        }

        [TestMethod()]
        public void PropertyBagContainsKeyHandlesLocalPropertyCacheTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            string noUpdateKey = _key + "_NoUpdate_" + DateTime.Now.ToFileTime();
            web.Context.ExecuteQueryRetry();

            props[_key] = _value_string;
            web.Update();

            props[noUpdateKey] = "This key is never added to the server because no web.Update() is called. This leads to future client caching issues if Web.AllProperties.RefreshLoad() or web.Context.Load(web.AllProperties) are called.";
            web.Context.ExecuteQueryRetry();

            Assert.IsTrue(web.PropertyBagContainsKey(_key));
            props.FieldValues.Remove(noUpdateKey); // Need to remove this key before refreshing from server

            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();

            Assert.IsFalse(props.FieldValues.ContainsKey(noUpdateKey), "The key '" + noUpdateKey + "' should not exist on the server");
        }


        [TestMethod()]
        public void GetIndexedPropertyBagKeysTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            var keys = web.GetIndexedPropertyBagKeys();

            Assert.IsInstanceOfType(keys, typeof(IEnumerable<string>), "No correct object returned");

            var keysList = keys.ToList();
            // Manually add an indexed property bag value
            if (!keysList.Contains(_key))
            {
                keysList.Add(_key);
                var encodedValues = GetEncodedValueForSearchIndexProperty(keysList);

                web.Context.Load(props);
                web.Context.ExecuteQueryRetry();

                props[INDEXED_PROPERTY_KEY] = encodedValues;

                web.Update();
                clientContext.ExecuteQueryRetry();
            }
            keys = web.GetIndexedPropertyBagKeys();
            Assert.IsTrue(keys.Contains(_key), "Key not present");

            // Local Cleanup
            props.RefreshLoad();
            clientContext.ExecuteQueryRetry();
            props[INDEXED_PROPERTY_KEY] = null;
            props.FieldValues.Remove(INDEXED_PROPERTY_KEY);
            web.Update();
            clientContext.ExecuteQueryRetry();
        }

        [TestMethod()]
        public void AddIndexedPropertyBagKeyTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;
            clientContext.Load(props);
            clientContext.ExecuteQueryRetry();

            web.AddIndexedPropertyBagKey(_key);

            props.RefreshLoad();
            clientContext.ExecuteQueryRetry();

            Assert.IsTrue(props.FieldValues.ContainsKey(INDEXED_PROPERTY_KEY));

            // Local cleanup
            props[INDEXED_PROPERTY_KEY] = null;
            props.FieldValues.Remove(INDEXED_PROPERTY_KEY);
            web.Update();
            clientContext.ExecuteQueryRetry();
        }

        [TestMethod()]
        public void RemoveIndexedPropertyBagKeyTest()
        {
            var web = clientContext.Web;
            var props = web.AllProperties;

            // Manually add an indexed property bag value
            var encodedValues = GetEncodedValueForSearchIndexProperty(new List<string>() { _key });

            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();

            props[INDEXED_PROPERTY_KEY] = encodedValues;

            web.Update();
            clientContext.ExecuteQueryRetry();

            // Remove the key
            Assert.IsTrue(web.RemoveIndexedPropertyBagKey(_key));
            props.RefreshLoad();
            clientContext.ExecuteQueryRetry();
            // All keys should be gone
            Assert.IsFalse(props.FieldValues.ContainsKey(_key), "Key still present");
        }
        #endregion

        #region ReIndex Tests
        [TestMethod()]
        public void TriggerReIndexTeamSiteTest()
        {
            var web = clientContext.Web;
            clientContext.Load(web);
            clientContext.ExecuteQueryRetry();
            web.ReIndexWeb();

            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();
            var version = (int)props["vti_searchversion"];
            web.ReIndexWeb();
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();
            var newVersion = (int)props["vti_searchversion"];
            Assert.IsTrue(version == (newVersion - 1), "Version has not increased");
        }
        #endregion

        #region Provisioning Tests
        [TestMethod]
        public void GetProvisioningTemplateTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var template = clientContext.Web.GetProvisioningTemplate();
                Assert.IsInstanceOfType(template, typeof(ProvisioningTemplate));
            }
        }

        [TestMethod]
        public void GetProvisioningTemplateWithSelectedContentTypesTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                // Arrange
                var creationInfo = new ProvisioningTemplateCreationInformation(web);
                creationInfo.ContentTypeGroupsToInclude.Add(contentTypeGroupName);
                creationInfo.HandlersToProcess = Handlers.ContentTypes;

                // Act
                var template = web.GetProvisioningTemplate(creationInfo);

                // Assert
                Assert.AreEqual(1, template.ContentTypes.Count);
                StringAssert.Equals(contentTypeGroupName, template.ContentTypes[0].Group);
            }
        }

        [TestMethod]
        public void GetProvisioningTemplateWithOutSelectedContentTypesTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                // Arrange
                var creationInfo = new ProvisioningTemplateCreationInformation(web);
                creationInfo.HandlersToProcess = Handlers.ContentTypes;

                // Act
                var template = web.GetProvisioningTemplate(creationInfo);

                // Assert
                Assert.IsTrue(template.ContentTypes.Count >= 1);
            }
        }
        #endregion

        #region NoScript tests
#if !ONPREMISES
        [TestMethod]
        public void IsNoScriptSiteTest()
        {
            if (String.IsNullOrEmpty(TestCommon.NoScriptSite))
            {
                Assert.Inconclusive("The NoScriptSite key was not set, test can't be executed.");
            }

            string devSiteUrl = TestCommon.DevSiteUrl;
            string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, TestCommon.NoScriptSite);

            using (var clientContext = TestCommon.CreateClientContext(siteToCreateUrl))
            {
                var isNoScriptSite = clientContext.Web.IsNoScriptSite();
                Assert.IsTrue(isNoScriptSite);

                isNoScriptSite = clientContext.Site.IsNoScriptSite();
                Assert.IsTrue(isNoScriptSite);
            }
        }

        [TestMethod]
        public void IsScriptSiteTest()
        {
            if (String.IsNullOrEmpty(TestCommon.ScriptSite))
            {
                Assert.Inconclusive("The ScriptSite key was not set, test can't be executed.");
            }

            string devSiteUrl = TestCommon.DevSiteUrl;
            string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, TestCommon.ScriptSite);

            using (var clientContext = TestCommon.CreateClientContext(siteToCreateUrl))
            {
                var isNoScriptSite = clientContext.Web.IsNoScriptSite();
                Assert.IsFalse(isNoScriptSite);

                isNoScriptSite = clientContext.Site.IsNoScriptSite();
                Assert.IsFalse(isNoScriptSite);
            }
        }


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
#endif
        #endregion

        #region App instance tests
        [TestMethod()]
        public void GetAppInstancesTest()
        {
            var web = clientContext.Web;

            var instances = web.GetAppInstances();
            Assert.IsInstanceOfType(instances, typeof(ClientObjectList<AppInstance>), "Incorrect return value");
            int instanceCount = instances.Count;

#if !ONPREMISES
            byte[] appToLoad = OfficeDevPnP.Core.Tests.Properties.Resources.HelloWorldApp;
#else
            byte[] appToLoad = OfficeDevPnP.Core.Tests.Properties.Resources.HelloWorldApp15;
#endif

            using (MemoryStream stream = new MemoryStream(appToLoad))
            {
                web.LoadApp(stream, 1033);
                clientContext.ExecuteQueryRetry();
            }

            instances = web.GetAppInstances();
            Assert.AreNotEqual(instances.Count, instanceCount, "App count is same after upload");
        }

        [TestMethod()]
        public void RemoveAppInstanceByTitleTest()
        {
            var web = clientContext.Web;

            var instances = web.GetAppInstances();
            Assert.IsInstanceOfType(instances, typeof(ClientObjectList<AppInstance>), "Incorrect return value");
            int instanceCount = instances.Count;

#if !ONPREMISES
            byte[] appToLoad = OfficeDevPnP.Core.Tests.Properties.Resources.HelloWorldApp;
#else
            byte[] appToLoad = OfficeDevPnP.Core.Tests.Properties.Resources.HelloWorldApp15;
#endif

            using (MemoryStream stream = new MemoryStream(appToLoad))
            {
                web.LoadApp(stream, 1033);
                clientContext.ExecuteQueryRetry();
            }

            string appToRemove = APPNAME;

#if ONPREMISES
            appToRemove += "15";
#endif

            Assert.IsTrue(web.RemoveAppInstanceByTitle(appToRemove));

            instances = web.GetAppInstances();

            Assert.AreEqual(instances.Count, instanceCount);
        }
        #endregion

        #region Install solution tests
        // DO NOT RUN. The DesignPackage.Install() function, used by this test, wipes the composed look gallery, breaking other tests.")]
        [Ignore()]
        [TestMethod()]
        public void InstallSolutionTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Set up

                // Write the test solution to a local temporary file
                string solutionpath = Path.Combine(Path.GetTempPath(), "testsolution.wsp");
                System.IO.File.WriteAllBytes(solutionpath, OfficeDevPnP.Core.Tests.Properties.Resources.TestSolution);

                clientContext.Site.InstallSolution(new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionGuid), solutionpath);

                // Check if the solution file is uploaded
                var solutionGallery = clientContext.Site.RootWeb.GetCatalog((int)ListTemplateType.SolutionCatalog);

                var camlQuery = new CamlQuery();
                camlQuery.ViewXml = string.Format(
      @"<View>  
            <Query> 
               <Where><Eq><FieldRef Name='SolutionId' /><Value Type='Guid'>{0}</Value></Eq></Where> 
            </Query> 
             <ViewFields><FieldRef Name='ID' /></ViewFields> 
      </View>", new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionGuid));

                var solutions = solutionGallery.GetItems(camlQuery);
                clientContext.Load(solutions);
                clientContext.ExecuteQueryRetry();

                // Test

                Assert.IsTrue(solutions.Any(), "No solution files available");

                // Check if we can activate Test Feature on rootweb
                clientContext.Load(clientContext.Web);
                clientContext.ExecuteQueryRetry();

                //  clientContext.Web.ActivateFeature(new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionFeatureGuid));
                //  Assert.IsTrue(clientContext.Web.IsFeatureActive(new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionFeatureGuid)), "Test feature not activated");

                // Teardown
                // Done using the local file, remove it
                System.IO.File.Delete(solutionpath);
                clientContext.Site.UninstallSolution(new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionGuid), "testsolution.wsp");
            }
        }

        // DO NOT RUN. The DesignPackage.Install() function, used by this test, wipes the composed look gallery, breaking other tests.")]
        [Ignore()]
        [TestMethod()]
        public void UninstallSolutionTest()
        {
            // Set up
            string solutionpath = Path.Combine(Path.GetTempPath(), "testsolution.wsp");
            System.IO.File.WriteAllBytes(solutionpath, OfficeDevPnP.Core.Tests.Properties.Resources.TestSolution);

            clientContext.Site.InstallSolution(new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionGuid), solutionpath);

            // Execute test

            clientContext.Site.UninstallSolution(new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionGuid), "testsolution.wsp");

            // Check if the solution file is uploaded
            var solutionGallery = clientContext.Site.RootWeb.GetCatalog((int)ListTemplateType.SolutionCatalog);

            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = string.Format(
  @"<View>  
            <Query> 
               <Where><Eq><FieldRef Name='SolutionId' /><Value Type='Guid'>{0}</Value></Eq></Where> 
            </Query> 
             <ViewFields><FieldRef Name='ID' /></ViewFields> 
      </View>", new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionGuid));

            var solutions = solutionGallery.GetItems(camlQuery);
            clientContext.Load(solutions);
            clientContext.ExecuteQueryRetry();
            Assert.IsFalse(solutions.Any(), "There are still solutions installed");

            Assert.IsFalse(clientContext.Web.IsFeatureActive(new Guid(OfficeDevPnP.Core.Tests.Properties.Resources.TestSolutionFeatureGuid)));

            // Teardown
            System.IO.File.Delete(solutionpath);
        }
        #endregion

        #region Various other tests
        [TestMethod]
        public void IsSubWebTest()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var site = ctx.Site;

                var rootWeb = site.RootWeb;

                var isSubweb = rootWeb.IsSubSite();

                Assert.IsFalse(isSubweb);
            }
        }
        #endregion

        #region ClientSide Package Deployment tests
#if !ONPREMISES
        [TestMethod()]
        public void DeploySharePointFrameworkSolutionTest()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {                
                var app = ctx.Web.DeployApplicationPackageToAppCatalog("hello-world.sppkg", "../../Resources", true, true, true);
            }
        }
#endif
        #endregion

        #region Helper methods
        private static string GetEncodedValueForSearchIndexProperty(IEnumerable<string> keys)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (string current in keys)
            {
                stringBuilder.Append(Convert.ToBase64String(Encoding.Unicode.GetBytes(current)));
                stringBuilder.Append('|');
            }
            return stringBuilder.ToString();
        }
        #endregion

        [TestMethod()]
        public void CanGetWebNameOfRootWebTest()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                string rootWebName = ctx.Web.GetName();
                Assert.AreEqual(string.Empty, rootWebName);
            }
        }

        [TestMethod()]
        public void CanGetWebNameOfSubWebTest()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                string subWebUrl = "Test_SubWebUrl";
                SiteEntity subSiteInfo = new SiteEntity()
                {
                    Title = "Test_Site Title",
                    Url = subWebUrl,
                    Template = "STS#0"
                };
                Web subSite = ctx.Web.CreateWeb(subSiteInfo);
                
                string subWebName = subSite.GetName();
                Assert.AreEqual(subWebUrl, subWebName);
            }
        }
    }
}
