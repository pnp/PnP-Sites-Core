using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class TokenParserTests
    {

        [TestMethod]
        public void ParseTests()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.Id, w => w.ServerRelativeUrl, w => w.Title, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                ctx.Load(ctx.Site, s => s.ServerRelativeUrl);

                var masterCatalog = ctx.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                ctx.Load(masterCatalog, m => m.RootFolder.ServerRelativeUrl);

                var themesCatalog = ctx.Web.GetCatalog((int)ListTemplateType.ThemeCatalog);
                ctx.Load(themesCatalog, t => t.RootFolder.ServerRelativeUrl);

                ctx.ExecuteQueryRetry();

                var currentUser = ctx.Web.EnsureProperty(w => w.CurrentUser);


                ProvisioningTemplate template = new ProvisioningTemplate();
                template.Parameters.Add("test", "test");

                var parser = new TokenParser(ctx.Web, template);
                var siteName = parser.ParseString("{sitename}");
                var siteId = parser.ParseString("{siteid}");
                var site1 = parser.ParseString("~siTE/test");
                var site2 = parser.ParseString("{site}/test");
                var sitecol1 = parser.ParseString("~siteCOLLECTION/test");
                var sitecol2 = parser.ParseString("{sitecollection}/test");
                var masterUrl1 = parser.ParseString("~masterpagecatalog/test");
                var masterUrl2 = parser.ParseString("{masterpagecatalog}/test");
                var themeUrl1 = parser.ParseString("~themecatalog/test");
                var themeUrl2 = parser.ParseString("{themecatalog}/test");
                var parameterTest1 = parser.ParseString("abc{parameter:TEST}/test");
                var parameterTest2 = parser.ParseString("abc{$test}/test");
                var associatedOwnerGroup = parser.ParseString("{associatedownergroup}");
                var associatedVisitorGroup = parser.ParseString("{associatedvisitorgroup}");
                var associatedMemberGroup = parser.ParseString("{associatedmembergroup}");
                var currentUserId = parser.ParseString("{currentuserid}");
                var currentUserLoginName = parser.ParseString("{currentuserloginname}");
                var currentUserFullName = parser.ParseString("{currentuserfullname}");
                var guid = parser.ParseString("{guid}");

                Assert.IsTrue(site1 == string.Format("{0}/test", ctx.Web.ServerRelativeUrl));
                Assert.IsTrue(site2 == string.Format("{0}/test", ctx.Web.ServerRelativeUrl));
                Assert.IsTrue(sitecol1 == string.Format("{0}/test", ctx.Site.ServerRelativeUrl));
                Assert.IsTrue(sitecol2 == string.Format("{0}/test", ctx.Site.ServerRelativeUrl));
                Assert.IsTrue(masterUrl1 == string.Format("{0}/test", masterCatalog.RootFolder.ServerRelativeUrl));
                Assert.IsTrue(masterUrl2 == string.Format("{0}/test", masterCatalog.RootFolder.ServerRelativeUrl));
                Assert.IsTrue(themeUrl1 == string.Format("{0}/test", themesCatalog.RootFolder.ServerRelativeUrl));
                Assert.IsTrue(themeUrl2 == string.Format("{0}/test", themesCatalog.RootFolder.ServerRelativeUrl));
                Assert.IsTrue(parameterTest1 == "abctest/test");
                Assert.IsTrue(parameterTest2 == "abctest/test");
                Assert.IsTrue(associatedOwnerGroup == ctx.Web.AssociatedOwnerGroup.Title);
                Assert.IsTrue(associatedVisitorGroup == ctx.Web.AssociatedVisitorGroup.Title);
                Assert.IsTrue(associatedMemberGroup == ctx.Web.AssociatedMemberGroup.Title);
                Assert.IsTrue(siteName == ctx.Web.Title);
                Assert.IsTrue(siteId == ctx.Web.Id.ToString());
                Assert.IsTrue(currentUserId == currentUser.Id.ToString());
                Assert.IsTrue(currentUserFullName == currentUser.Title);
                Assert.IsTrue(currentUserLoginName == currentUser.LoginName);
                Guid outGuid;
                Assert.IsTrue(Guid.TryParse(guid, out outGuid));
            }
        }
    }
}
