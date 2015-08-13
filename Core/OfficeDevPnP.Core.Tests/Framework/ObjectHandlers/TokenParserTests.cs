using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Configuration;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

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
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.Load(ctx.Site, s => s.ServerRelativeUrl);

                var masterCatalog = ctx.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                ctx.Load(masterCatalog, m => m.RootFolder.ServerRelativeUrl);

                var themesCatalog = ctx.Web.GetCatalog((int)ListTemplateType.ThemeCatalog);
                ctx.Load(themesCatalog, t => t.RootFolder.ServerRelativeUrl);

                ctx.ExecuteQueryRetry();

                
                ProvisioningTemplate template = new ProvisioningTemplate();
                template.Parameters.Add("test", "test");

                var parser = new TokenParser(ctx.Web, template);

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

            }
        }
    }
}
