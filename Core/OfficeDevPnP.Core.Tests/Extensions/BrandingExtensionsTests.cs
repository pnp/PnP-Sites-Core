using System.Security.Cryptography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Configuration;
using OfficeDevPnP.Core.Utilities;
using System.Diagnostics;

namespace OfficeDevPnP.Core.Tests.AppModelExtensions
{
    [TestClass()]
    public class BrandingExtensionsTests
    {
        private string builtInLookSeaMonster = "Sea Monster"; // oslo, palette005, image_bg005, fontscheme003
        private string builtInLookBlossom = "Blossom"; // seattle, palette002,image_bg002
        private string builtInMasterOslo = "oslo.master";
        private string builtInMasterSeattle = "seattle.master";
        private string builtInPalette003 = "palette003.spcolor";
        private string knownHashOfSeattle = "DA-39-A3-EE-5E-6B-4B-0D-32-55-BF-EF-95-60-18-90-AF-D8-07-09";

        private string customColorFilePath = string.Empty;
        private string customBackgroundFilePath = string.Empty;
        private const string THEME_NAME = "Test_Theme";

        const string CAML_QUERY_FIND_BY_FILENAME = @"
                <View>
                    <Query>                
                        <Where>
                            <Eq>
                                <FieldRef Name='Name' />
                                <Value Type='Text'>{0}</Value>
                            </Eq>
                        </Where>
                     </Query>
                </View>";
        
        private static string htmlPublishingPageWithoutExtension = "TestHtmlPublishingPageLayout";
        private static string publishingPageWithoutExtension = "TestPublishingPageLayout";
        private string htmlPublishingPagePath = string.Format("../../Resources/{0}.html", htmlPublishingPageWithoutExtension);
        private string publishingPagePath = string.Format("../../Resources/{0}.aspx", publishingPageWithoutExtension);
        private string pageLayoutTitle = "CustomHtmlPageLayout";
        private string welcomePageContentTypeId = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF390064DEA0F50FC8C147B0B6EA0636C4A7D4";
        private string testWebName;
        bool deactivateSiteFeatureOnTeardown = false;
        bool deactivateWebFeatureOnTeardown = false;
        private Web pageLayoutTestWeb = null;
        private string AvailablePageLayouts = "__PageLayouts";

        #region Test initialize and cleanup
        [TestInitialize()]
        public void Initialize()
        {
            Console.WriteLine("BrandingExtensionsTests.Initialise");

            customColorFilePath = Path.Combine(Path.GetTempPath(), "custom.spcolor");
            System.IO.File.WriteAllBytes(customColorFilePath, OfficeDevPnP.Core.Tests.Properties.Resources.custom);
            customBackgroundFilePath = Path.Combine(Path.GetTempPath(), "custombg.jpg");
            //Properties.Resources.custombg.Save(customBackgroundFilePath);

            testWebName = string.Format("Test_CL{0:yyyyMMddTHHmmss}", DateTimeOffset.Now);
            pageLayoutTestWeb = Setup();
        }

        [TestCleanup()]
        public void CleanUp()
        {
            Console.WriteLine("BrandingExtensionsTests.CleanUp");

            if (System.IO.File.Exists(customColorFilePath))
            {
                System.IO.File.Delete(customColorFilePath);
            }
            if (System.IO.File.Exists(customBackgroundFilePath))
            {
                System.IO.File.Delete(customBackgroundFilePath);
            }

            using (var context = TestCommon.CreateClientContext())
            {
                var web = context.Web;

                // Remove composed looks from server
                List themeGallery = web.GetCatalog((int)ListTemplateType.DesignCatalog);
                CamlQuery query = new CamlQuery();
                string camlString = @"
                    <View>
                        <Query>                
                            <Where>
                                <Contains>
                                    <FieldRef Name='Name' />
                                    <Value Type='Text'>Test_</Value>
                                </Contains>
                            </Where>
                        </Query>
                    </View>";
                query.ViewXml = camlString;
                var found = themeGallery.GetItems(query);
                web.Context.Load(found);
                web.Context.ExecuteQueryRetry();
                Console.WriteLine("{0} matching looks found to delete", found.Count);
                var looksToDelete = found.ToList();
                foreach (var item in looksToDelete)
                {
                    Console.WriteLine("Delete look item '{0}'", item["Name"]);
                    item.DeleteObject();
                    context.ExecuteQueryRetry();
                }

                // Remove Theme Files
                List themesList = web.GetCatalog((int)ListTemplateType.ThemeCatalog);
                Folder rootFolder = themesList.RootFolder;
                FolderCollection rootFolders = rootFolder.Folders;
                web.Context.Load(rootFolder);
                web.Context.Load(rootFolders, f => f.Where(folder => folder.Name == "15"));
                web.Context.ExecuteQueryRetry();

                Folder folder15 = rootFolders.FirstOrDefault();

                try
                {
                    Microsoft.SharePoint.Client.File customColorFile = folder15.Files.GetByUrl("custom.spcolor");
                    customColorFile.DeleteObject();
                    context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception cleaning up: {0}", ex);
                }

                try
                {
                    Microsoft.SharePoint.Client.File customBackgroundFile = folder15.Files.GetByUrl("custombg.jpg");
                    customBackgroundFile.DeleteObject();
                    context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception cleaning up: {0}", ex);
                }

                // Remove webs
                var webCollection1 = web.Webs;
                context.Load(webCollection1, wc => wc.Include(w => w.Title, w => w.ServerRelativeUrl));
                context.ExecuteQueryRetry();
                var websToDelete = new List<Web>();
                foreach (var web1 in webCollection1)
                {
                    if (web1.Title.StartsWith("Test_"))
                    {
                        var webCollection2 = web1.Webs;
                        context.Load(webCollection2, wc => wc.Include(w => w.Title, w => w.ServerRelativeUrl));
                        context.ExecuteQueryRetry();
                        var childrenToDelete = new List<Web>(webCollection2);
                        foreach (var web2 in childrenToDelete)
                        {
                            Console.WriteLine("Deleting site {0}", web2.ServerRelativeUrl);
                            web2.DeleteObject();
                            context.ExecuteQueryRetry();
                        }
                        websToDelete.Add(web1);
                    }
                }

                foreach (var web1 in websToDelete)
                {
                    Console.WriteLine("Deleting site {0}", web1.ServerRelativeUrl);
                    web1.DeleteObject();
                    try
                    {
                        context.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Exception cleaning up: {0}", ex);
                    }
                }

                // Remove pagelayouts
                List masterPageGallery = context.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                Folder rootFolderInMasterPageGallery = masterPageGallery.RootFolder;
                context.Load(rootFolderInMasterPageGallery, f => f.ServerRelativeUrl);
                context.ExecuteQueryRetry();

                try
                {
                    var fileServerRelativeUrl = UrlUtility.Combine(rootFolderInMasterPageGallery.ServerRelativeUrl, publishingPageWithoutExtension);
                    var file = context.Web.GetFileByServerRelativeUrl(String.Format("{0}.aspx", fileServerRelativeUrl));
                    context.Load(file);
                    context.ExecuteQueryRetry();
                    file.DeleteObject();
                    context.ExecuteQueryRetry();

                    fileServerRelativeUrl = UrlUtility.Combine(rootFolderInMasterPageGallery.ServerRelativeUrl, "test/test", publishingPageWithoutExtension);
                    file = context.Web.GetFileByServerRelativeUrl(String.Format("{0}.aspx", fileServerRelativeUrl));
                    context.Load(file);
                    context.ExecuteQueryRetry();
                    file.DeleteObject();
                    context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception cleaning up: {0}", ex);
                }
            }

            Teardown();
        }

        private Web Setup()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                Web web;
                Site site;
                site = ctx.Site;
                web = ctx.Site.RootWeb;

                if (!site.IsFeatureActive(Constants.FeatureId_Site_Publishing))
                {
                    site.ActivateFeature(Constants.FeatureId_Site_Publishing);
                    Console.WriteLine("Site publishing feature activated");
                    deactivateSiteFeatureOnTeardown = true;
                }

                if (!web.IsFeatureActive(Constants.FeatureId_Web_Publishing))
                {
                    site.RootWeb.ActivateFeature(Constants.FeatureId_Web_Publishing);
                    Console.WriteLine("Web publishing feature activated");
                    deactivateWebFeatureOnTeardown = true;
                }

                using (var wciCtx = ctx.Clone(TestCommon.DevSiteUrl))
                {
                    CreatePublishingWeb(wciCtx, testWebName);
                }

                using (var web1Ctx = ctx.Clone(TestCommon.DevSiteUrl + "/" + testWebName))
                {
                    web1Ctx.Web.ActivateFeature(new Guid("41E1D4BF-B1A2-47F7-AB80-D5D6CBBA3092"));
                    Console.WriteLine("Feature 41E1D4BF-B1A2-47F7-AB80-D5D6CBBA3092 on Web {0} activated", testWebName);
                }

                using (var web1Ctx = ctx.Clone(TestCommon.DevSiteUrl + "/" + testWebName))
                {
                    CreatePublishingWeb(web1Ctx, "A");

                    using (var webACtx = ctx.Clone(TestCommon.DevSiteUrl + "/" + testWebName + "/a"))
                    {
                        webACtx.Web.ActivateFeature(new Guid("41E1D4BF-B1A2-47F7-AB80-D5D6CBBA3092"));
                        Console.WriteLine("Feature 41E1D4BF-B1A2-47F7-AB80-D5D6CBBA3092 on Web {0} activated", "a");
                    }
                }

                using (var web1Ctx = ctx.Clone(TestCommon.DevSiteUrl + "/" + testWebName))
                {
                    CreatePublishingWeb(web1Ctx, "B");

                    using (var webBCtx = ctx.Clone(TestCommon.DevSiteUrl + "/" + testWebName + "/b"))
                    {
                        webBCtx.Web.ActivateFeature(new Guid("41E1D4BF-B1A2-47F7-AB80-D5D6CBBA3092"));
                        Console.WriteLine("Feature 41E1D4BF-B1A2-47F7-AB80-D5D6CBBA3092 on Web {0} activated", "b");
                    }
                }

                return web;
            }
        }

        private void CreatePublishingWeb(ClientContext wciCtx, string urlAndName)
        {
            int retryAttempts = 0;
            int retryCount = 10;
            int delay = 500;
            int backoffInterval = delay;

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
                {
                    var wci1 = new WebCreationInformation();
                    wci1.Url = urlAndName;
                    wci1.Title = urlAndName;
                    wci1.WebTemplate = "CMSPUBLISHING#0";
                    var web1 = wciCtx.Web.Webs.Add(wci1);
                    wciCtx.ExecuteQueryRetry();
                    Console.WriteLine("Web {0} created", urlAndName);
                    return;
                }
                catch (Microsoft.SharePoint.Client.ServerException ex)
                {
                    Console.WriteLine(ex.ToDetailedString(wciCtx));
                    //retry
                    Console.WriteLine("Site creation failed. Sleeping for {0} seconds before retrying.", backoffInterval);

                    //Add delay for retry
                    System.Threading.Thread.Sleep(backoffInterval);

                    //Add to retry count and increase delay.
                    retryAttempts++;
                    backoffInterval = backoffInterval * 2;
                }
            }
            throw new Exception(string.Format("Maximum site creation retry attempts {0} has been reached.", retryCount));
        }

        private void Teardown()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                if (deactivateSiteFeatureOnTeardown)
                {
                    ctx.Site.DeactivateFeature(Constants.FeatureId_Site_Publishing);
                }
                if (deactivateWebFeatureOnTeardown)
                {
                    ctx.Web.DeactivateFeature(Constants.FeatureId_Web_Publishing);
                }
            }
        }
        #endregion

        #region Pagelayout tests
        [TestMethod]
        public void CanUploadHtmlPageLayoutAndConvertItToAspxVersionTest()
        {
            var web = pageLayoutTestWeb;
            web.Context.Load(web);

            // First delete the published file (if any) to avoid running into to many minor versions error with automated testing
            List masterPageGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            Folder rootFolder = masterPageGallery.RootFolder;
            web.Context.Load(masterPageGallery);
            web.Context.Load(rootFolder);
            web.Context.ExecuteQueryRetry();

            Microsoft.SharePoint.Client.File htmlPublishLayout = rootFolder.GetFile(String.Format("{0}.html", htmlPublishingPageWithoutExtension));
            if (htmlPublishLayout != null)
            {
                htmlPublishLayout.DeleteObject();
                web.Context.ExecuteQueryRetry();
            }
            
            web.DeployHtmlPageLayout(htmlPublishingPagePath, pageLayoutTitle, "", welcomePageContentTypeId);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            var item = web.GetPageLayoutListItemByName(htmlPublishingPageWithoutExtension);
            Assert.AreNotEqual(null, item);
        }

        [TestMethod]
        public void CanUploadPageLayoutTest()
        {
            var web = pageLayoutTestWeb;
            web.Context.Load(web);
            web.DeployPageLayout(publishingPagePath, pageLayoutTitle, "", welcomePageContentTypeId);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            var item = web.GetPageLayoutListItemByName(publishingPageWithoutExtension);
            Assert.AreNotEqual(null, item);
        }

        [TestMethod]
        public void CanUploadPageLayoutWithPathTest()
        {
            var web = pageLayoutTestWeb;
            web.Context.Load(web);
            web.DeployPageLayout(publishingPagePath, pageLayoutTitle, "", welcomePageContentTypeId, "test/test");
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            var item = web.GetPageLayoutListItemByName("test/test/" + publishingPageWithoutExtension);
            Assert.AreNotEqual(null, item);
        }

        [TestMethod]
        public void AllowAllPageLayoutsTest()
        {
            var web = pageLayoutTestWeb;

            web.AllowAllPageLayouts();

            string allowedPageLayouts = web.GetPropertyBagValueString(AvailablePageLayouts, null);

            Assert.AreEqual(allowedPageLayouts, string.Empty);
        }

        #endregion

        #region Composed Look tests
        [TestMethod()]
        public void DeployThemeAndCreateComposedLookTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                context.Web.UploadThemeFile(customColorFilePath);
                context.Web.UploadThemeFile(customBackgroundFilePath);
                context.Web.CreateComposedLookByName("Test_Theme", customColorFilePath, null, customBackgroundFilePath, null);
                Assert.IsTrue(context.Web.ComposedLookExists("Test_Theme"));
            }
        }

        [TestMethod()]
        public void ComposedLookExistsTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                Assert.IsTrue(context.Web.ComposedLookExists("Office"));
                Assert.IsFalse(context.Web.ComposedLookExists("Dummy Test Theme That Should Not Exist"));
            }
        }

        [TestMethod()]
        public void GetCurrentComposedLookTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                context.Web.SetComposedLookByUrl(builtInLookSeaMonster);
            }

            using (var context = TestCommon.CreateClientContext())
            {
                var theme = context.Web.GetCurrentComposedLook();
                Assert.IsTrue(theme != null);
                Assert.IsTrue(theme.BackgroundImage.EndsWith("image_bg005.jpg"));
            }
        }

        [TestMethod()]
        public void CreateComposedLookShouldWorkTest()
        {
            var testLookName = string.Format("Test_CL{0:yyyyMMddTHHmmss}", DateTimeOffset.Now);

            using (var context = TestCommon.CreateClientContext())
            {
                context.Load(context.Web, w => w.ServerRelativeUrl);
                context.ExecuteQueryRetry();
                var paletteServerRelativeUrl = context.Web.ServerRelativeUrl + "/_catalog/theme/15" + builtInPalette003;
                var masterServerRelativeUrl = context.Web.ServerRelativeUrl + "/_catalog/masterpage" + builtInMasterOslo;

                context.Web.CreateComposedLookByUrl(testLookName, paletteServerRelativeUrl, null, null, masterServerRelativeUrl, 5);
            }

            using (var context = TestCommon.CreateClientContext())
            {
                var composedLooksList = context.Web.GetCatalog((int)ListTemplateType.DesignCatalog);
                CamlQuery query = new CamlQuery();
                query.ViewXml = string.Format(CAML_QUERY_FIND_BY_FILENAME, testLookName);
                var existingCollection = composedLooksList.GetItems(query);
                context.Load(existingCollection);
                context.ExecuteQueryRetry();
                var item = existingCollection.FirstOrDefault();

                var lookPaletteUrl = item["ThemeUrl"] as FieldUrlValue;
                Assert.IsTrue(lookPaletteUrl.Url.Contains(builtInPalette003));
                var lookMasterUrl = item["MasterPageUrl"] as FieldUrlValue;
                Assert.IsTrue(lookMasterUrl.Url.Contains(builtInMasterOslo));
                var lookDisplayOrder = item["DisplayOrder"].ToString();
                Assert.AreEqual("5", lookDisplayOrder);
            }
        }

        [TestMethod()]
        public void CreateComposedLookByNameShouldWorkTest()
        {
            var testLookName = string.Format("Test_CL{0:yyyyMMddTHHmmss}", DateTimeOffset.Now);

            using (var context = TestCommon.CreateClientContext())
            {
                // Act
                context.Web.CreateComposedLookByName(testLookName, builtInPalette003, null, null, builtInMasterOslo, 5);
            }

            using (var context = TestCommon.CreateClientContext())
            {
                var composedLooksList = context.Web.GetCatalog((int)ListTemplateType.DesignCatalog);
                CamlQuery query = new CamlQuery();
                query.ViewXml = string.Format(CAML_QUERY_FIND_BY_FILENAME, testLookName);
                var existingCollection = composedLooksList.GetItems(query);
                context.Load(existingCollection);
                context.ExecuteQueryRetry();
                var item = existingCollection.FirstOrDefault();

                var lookPaletteUrl = item["ThemeUrl"] as FieldUrlValue;
                Assert.IsTrue(lookPaletteUrl.Url.Contains(builtInPalette003));
                var lookMasterUrl = item["MasterPageUrl"] as FieldUrlValue;
                Assert.IsTrue(lookMasterUrl.Url.Contains(builtInMasterOslo));
            }
        }

        [TestMethod()]
        public void SetComposedLookInheritsTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                var webCollection = context.Web.Webs;
                context.Load(webCollection, wc => wc.Include(w => w.Title));
                context.ExecuteQueryRetry();
                var webToChange1 = webCollection.First(w => w.Title == testWebName);

                var webCollection2 = webToChange1.Webs;
                context.Load(webCollection2, wc => wc.Include(w => w.Title));
                context.ExecuteQueryRetry();
                var webToChangeA = webCollection2.First(w => w.Title == "A");

                // Act
                webToChangeA.SetComposedLookByUrl(builtInLookBlossom);
                webToChange1.SetComposedLookByUrl(builtInLookSeaMonster, resetSubsitesToInherit: false);
            }

            using (var context = TestCommon.CreateClientContext())
            {
                var webCollection = context.Web.Webs;
                context.Load(webCollection, wc => wc.Include(w => w.Title));
                context.ExecuteQueryRetry();
                var webToCheck1 = webCollection.First(w => w.Title == testWebName);

                var webCollection2 = webToCheck1.Webs;
                context.Load(webCollection2, wc => wc.Include(w => w.Title, w => w.MasterUrl, w => w.CustomMasterUrl));
                context.ExecuteQueryRetry();

                var webToCheckB = webCollection2.First(w => w.Title == "B");
                var webToCheckA = webCollection2.First(w => w.Title == "A");
                var accentTextB = webToCheckB.ThemeInfo.GetThemeShadeByName("AccentText");
                var accentTextA = webToCheckA.ThemeInfo.GetThemeShadeByName("AccentText");
                context.ExecuteQueryRetry();

                // Assert: B will have new style, A will have Inherit = false and not get the new style

                // Sea Monster oslo, palette005, image_bg005, fontscheme003
                Assert.IsTrue(webToCheckB.MasterUrl.Contains(builtInMasterOslo));
                Assert.AreEqual("FFF07200", accentTextB.Value);

                // Blossom seattle, palette002, image_bg002
                Assert.IsTrue(webToCheckA.MasterUrl.Contains(builtInMasterSeattle));
                Assert.AreEqual("FFD55881", accentTextA.Value);
            }
        }

        [TestMethod()]
        public void SetComposedLookResetInheritanceTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                var webCollection = context.Web.Webs;
                context.Load(webCollection, wc => wc.Include(w => w.Title));
                context.ExecuteQueryRetry();
                var webToChangeRoot = webCollection.First(w => w.Title == testWebName);

                var webCollection2 = webToChangeRoot.Webs;
                context.Load(webCollection2, wc => wc.Include(w => w.Title));
                context.ExecuteQueryRetry();
                var webToChangeA = webCollection2.First(w => w.Title == "A");

                // Act
                webToChangeA.SetComposedLookByUrl(builtInLookBlossom);
                webToChangeRoot.SetComposedLookByUrl(builtInLookSeaMonster, resetSubsitesToInherit: true, updateRootOnly:false);
            }

            using (var context = TestCommon.CreateClientContext())
            {
                var webCollection = context.Web.Webs;
                context.Load(webCollection, wc => wc.Include(w => w.Title));
                context.ExecuteQueryRetry();
                var webToCheckRoot = webCollection.First(w => w.Title == testWebName);

                var webCollection2 = webToCheckRoot.Webs;
                context.Load(webCollection2, wc => wc.Include(w => w.Title, w => w.MasterUrl, w => w.CustomMasterUrl));
                context.ExecuteQueryRetry();
                var webToCheckA = webCollection2.First(w => w.Title == "A");
                var accentA = webToCheckA.ThemeInfo.GetThemeShadeByName("AccentText");
                context.ExecuteQueryRetry();

                // Assert: B will have Inherit = false and not get the new style, A will hav new style

                // Sea Monster oslo, palette005, image_bg005, fontscheme003
                Assert.IsTrue(webToCheckA.MasterUrl.Contains(builtInMasterOslo));
                Assert.AreEqual("FFF07200", accentA.Value);
            }
        }
        #endregion

        #region Master page tests
        // Manually taken over from Gavin Barron's commit https://github.com/gavinbarron/PnP/blob/17c4d3647f4a509fb1eedb949ef07af7f962929c/OfficeDevPnP.Core/OfficeDevPnP.Core.Tests/AppModelExtensions/BrandingExtensionsTests.cs 
        [TestMethod]
        public void SeattleMasterPageIsUnchangedTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                var web = context.Web;
                //need to get the server relative url 
                context.Load(web, w => w.ServerRelativeUrl);
                context.ExecuteQueryRetry();
                //Use the existing context to directly get a copy of the seattle master page 
                string masterpageGalleryServerRelativeUrl = UrlUtility.Combine(UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl), "_catalogs/masterpage/");
                var serverRelativeUrlOfSeattle = UrlUtility.Combine(masterpageGalleryServerRelativeUrl, builtInMasterSeattle);

                // OpenBinaryDirect fails when used with app only
                //FileInformation seattle = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, serverRelativeUrlOfSeattle);
                var seattle = context.Web.GetFileByServerRelativeUrl(serverRelativeUrlOfSeattle);
                web.Context.Load(seattle);
                web.Context.ExecuteQueryRetry();

                Assert.IsNotNull(seattle);

                ClientResult<Stream> data = seattle.OpenBinaryStream();
                context.Load(seattle);
                context.ExecuteQueryRetry();

                //Dump seattle.master
                //if (data != null)
                //{
                //    int position = 1;
                //    int bufferSize = 200000;
                //    Byte[] readBuffer = new Byte[bufferSize];
                //    string localFilePath = "C:\\Temp\\seattle.master";
                //    using (System.IO.Stream stream = System.IO.File.Create(localFilePath))
                //    {
                //        while (position > 0)
                //        {
                //            // data.Value holds the Stream
                //            position = data.Value.Read(readBuffer, 0, bufferSize);
                //            stream.Write(readBuffer, 0, position);
                //            readBuffer = new Byte[bufferSize];
                //        }
                //        stream.Flush();
                //    }
                //}

                MemoryStream memStream = new MemoryStream();
                data.Value.CopyTo(memStream);

                //Compute a hash of the file 
                var hashAlgorithm = HashAlgorithm.Create();
                byte[] hash = hashAlgorithm.ComputeHash(memStream);
                //Convert to a hex string for human consumption 
                string hex = BitConverter.ToString(hash);
                //Check against last known hash 
                Assert.AreEqual(knownHashOfSeattle, hex);
            }
        }
        #endregion

        #region Miscellanious tests
        [TestMethod]
        public void IsSubsiteTest()
        {
            using (ClientContext cc = TestCommon.CreateClientContext())
            {
                Assert.IsFalse(cc.Web.IsSubSite());

                using (ClientContext ctx = cc.Clone(string.Format("{0}/{1}", ConfigurationManager.AppSettings["SPODevSiteUrl"], testWebName)))
                {
                    Assert.IsTrue(ctx.Web.IsSubSite());
                }
            }
        }
        #endregion

        #region Responsive tests
        [TestMethod]
        public void CanEnableResponsiveDesignTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                context.Web.EnableResponsiveUI();
            }
        }

        [TestMethod]
        public void CanDisableResponsiveDesignTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                context.Web.DisableReponsiveUI();
            }
        }
        #endregion
    }
}
