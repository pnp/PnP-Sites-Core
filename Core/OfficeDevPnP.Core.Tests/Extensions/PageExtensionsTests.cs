using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Tests.AppModelExtensions
{
    [TestClass]
    public class PageExtensionsTests
    {
        private string folder = "SitePages";
        private string pageName = "Home.aspx";
        private string publishingPageName = "Happy";
        private string publishingPageTemplate = "BlankWebPartPage";
        private Guid publishingSiteFeatureId = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");
        bool deactivatePublishingOnTearDown = false;

        #region Test initialize and cleanup
        public Web Setup(string webTemplate = "STS#0", bool enablePublishingInfrastructure = false)
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var name = "WebExtensions";
                ctx.ExecuteQueryRetry();

                Web web;
                Site site;
                site = ctx.Site;
                if (enablePublishingInfrastructure && !site.IsFeatureActive(publishingSiteFeatureId))
                {
                    site.ActivateFeature(publishingSiteFeatureId);
                    deactivatePublishingOnTearDown = true;
                }

                try
                {
                    web = ctx.Site.OpenWeb(name);
                    web.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
                catch { }

                using (var webCtx = ctx.Clone(TestCommon.DevSiteUrl))
                {
                    web = webCtx.Web.Webs.Add(new WebCreationInformation
                    {
                        Title = name,
                        WebTemplate = webTemplate,
                        Url = name
                    });
                    webCtx.ExecuteQueryRetry();
                }

                // Wait for 2 minutes after creating a publishing web...especially when using app-only as otherwise one 
                // can get a "Microsoft.SharePoint.Client.ServerException: The site is not valid. The 'Pages' document library is missing." error
                // during subsequent creation/deletion actions
                if (webTemplate.Equals("CMSPUBLISHING#0", StringComparison.InvariantCultureIgnoreCase))
                {
                    System.Threading.Thread.Sleep(3 * 60 * 1000);
                }

                // Create client context object for the newly created web and return that one...avoids "request uses too many resources" errors
                using (var newWebCtx = ctx.Clone(TestCommon.DevSiteUrl + "/" + name))
                {
                    return newWebCtx.Web;
                }
            }
        }

        public void Teardown(Web web)
        {
            web.DeleteObject();
            if(deactivatePublishingOnTearDown)
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    ctx.Site.DeactivateFeature(publishingSiteFeatureId);
                }
            }
        }
        #endregion

        #region Wiki page tests
        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void AddWikiPageTest()
        {
            var web = Setup();

            //Add new wiki library
            web.CreateList(ListTemplateType.WebPageLibrary, "wikipages", false);

            using (var ctx = TestCommon.CreateClientContext())
            {
                using (var ctx2 = ctx.Clone(TestCommon.DevSiteUrl + "/WebExtensions"))
                {
                    var wikiPage2 = ctx2.Web.AddWikiPage("wikipages", "test.aspx");
                    Assert.AreEqual(wikiPage2.ToLower(), "wikipages/test.aspx");
                    var wikiPage1 = ctx2.Web.AddWikiPage("Site Pages", "test.aspx");
                    Assert.AreEqual(wikiPage1.ToLower(), "sitepages/test.aspx");

                    Teardown(ctx2.Web);
                }
            }
        }

        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void EnsureWikiPageTest()
        {
            var web = Setup();

            using (var ctx = TestCommon.CreateClientContext())
            {
                using (var ctx2 = ctx.Clone(TestCommon.DevSiteUrl + "/WebExtensions"))
                {
                    // first run creates the page
                    var wikiPage1 = ctx2.Web.EnsureWikiPage("Site Pages", "test.aspx");
                    Assert.AreEqual(wikiPage1.ToLower(), "sitepages/test.aspx");
                    
                    // Second run should return the page
                    wikiPage1 = ctx2.Web.EnsureWikiPage("Site Pages", "test.aspx");
                    Assert.AreEqual(wikiPage1.ToLower(), "sitepages/test.aspx");

                    Teardown(ctx2.Web);
                }
            }
        }

        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void CanAddLayoutToWikiPageTest()
        {
            var web = Setup();

            web.AddLayoutToWikiPage(folder, OfficeDevPnP.Core.WikiPageLayout.TwoColumns, pageName);

            Teardown(web);
        }

	    [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void CanAddHtmlToWikiPageTest()
        {
            var web = Setup();

            web.AddHtmlToWikiPage(folder, "<h1>I got text</h1>", pageName, 1, 1);

            Teardown(web);
        }

        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void ProveThatWeCanAddHtmlToPageAfterChangingLayoutTest()
        {
            var web = Setup();
            web.AddLayoutToWikiPage(folder, OfficeDevPnP.Core.WikiPageLayout.TwoColumns, pageName);
            web.AddHtmlToWikiPage(folder, "<h1>I got text</h1>", pageName, 1, 1);

            var content = web.GetWikiPageContent(UrlUtility.Combine(UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl), folder, pageName));

            Assert.IsTrue(content.Contains("<h1>I got text</h1>"));

            Teardown(web);
        }
        #endregion

        #region Publishing page tests
        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void CanCreatePublishingPageTest()
        {
            var web = Setup("CMSPUBLISHING#0",true);
            web.Context.Load(web);
            web.AddPublishingPage(publishingPageName, publishingPageTemplate);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            var page = web.GetPublishingPage(string.Format("{0}.aspx",publishingPageName));
            web.Context.Load(page.ListItem, i => i["Title"]);
            web.Context.ExecuteQueryRetry();
            
            Assert.AreEqual(page.ListItem["Title"], publishingPageName);

            Teardown(web);
        }

        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void PublishingPageWithInvalidCharsIsCorrectlyCreatedTest()
        {
            var web = Setup("CMSPUBLISHING#0", true);
            web.Context.Load(web);
            web.AddPublishingPage("Happy?is:good", publishingPageTemplate);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            var page = web.GetPublishingPage(string.Format("{0}.aspx", "Happy-is-good"));
            Assert.IsNotNull(page);

            Teardown(web);
        }

        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void CanCreatePublishedPublishingPageWhenModerationIsEnabledTest()
        {
            var web = Setup("CMSPUBLISHING#0", true);
            web.Context.Load(web);
            //Ensure that moderation is enabled
            var pagesLibrary = web.Lists.GetByTitle("Pages");
            pagesLibrary.EnableModeration = true;
            pagesLibrary.Update();
            web.Context.ExecuteQueryRetry();
            web.AddPublishingPage(publishingPageName, publishingPageTemplate,publish:true);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            var page = web.GetPublishingPage(string.Format("{0}.aspx", publishingPageName));
            web.Context.Load(page.ListItem, i => i["_ModerationStatus"]);
            web.Context.Load(page.ListItem, i => i.File.MajorVersion);
            web.Context.ExecuteQueryRetry();

            Assert.AreEqual(0, page.ListItem["_ModerationStatus"]);
            Assert.AreEqual(1, page.ListItem.File.MajorVersion);

            Teardown(web);
        }

        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void CanCreatePublishedPublishingPageWhenModerationIsDisabledTest()
        {
            var web = Setup("CMSPUBLISHING#0", true);
            web.Context.Load(web);
            //Ensure that moderation is disabled
            var pagesLibrary = web.Lists.GetByTitle("Pages");
            pagesLibrary.EnableModeration = false;
            pagesLibrary.Update();
            web.Context.ExecuteQueryRetry();
            web.AddPublishingPage(publishingPageName, publishingPageTemplate, publish: true);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            var page = web.GetPublishingPage(string.Format("{0}.aspx", publishingPageName));
            web.Context.Load(page.ListItem, i => i.File.MajorVersion);
            web.Context.ExecuteQueryRetry();

            Assert.AreEqual(1, page.ListItem.File.MajorVersion);

            Teardown(web);
        }

        [TestMethod]
        [Timeout(10 * 60 * 1000)]
        public void CreatedPublishingPagesSetsTitleCorrectlyTest()
        {
            var web = Setup("CMSPUBLISHING#0", true);
            var customTitle = "Bad robot";
            var customPublishingPageName = "Good-robot";
            web.Context.Load(web);
            //Ensure that moderation is enabled
            var pagesLibrary = web.Lists.GetByTitle("Pages");
            pagesLibrary.EnableModeration = true;
            pagesLibrary.Update();
            web.Context.ExecuteQueryRetry();
            web.AddPublishingPage(publishingPageName, publishingPageTemplate, publish: true);
            web.AddPublishingPage(customPublishingPageName, publishingPageTemplate, publish: true, title: customTitle);
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            var pageWithNoTitle = web.GetPublishingPage(string.Format("{0}.aspx", publishingPageName));
            var pageWithCustomTitle = web.GetPublishingPage(string.Format("{0}.aspx", customPublishingPageName));
            web.Context.Load(pageWithNoTitle.ListItem, i => i["Title"]);
            web.Context.Load(pageWithCustomTitle.ListItem, i => i["Title"]);
            web.Context.ExecuteQueryRetry();

            //Check that title is set to page name
            Assert.AreEqual(publishingPageName, pageWithNoTitle.ListItem["Title"]);
            //Check that title is set to title
            Assert.AreEqual(customTitle, pageWithCustomTitle.ListItem["Title"]);

            Teardown(web);
        }
        #endregion
    }
}
