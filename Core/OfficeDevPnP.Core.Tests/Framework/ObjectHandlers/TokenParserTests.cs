using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
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
                ctx.Load(ctx.Web, 
                    w => w.Id, 
                    w => w.ServerRelativeUrl, 
                    w => w.Title,
                    w => w.AssociatedOwnerGroup.Title, 
                    w => w.AssociatedMemberGroup.Title, 
                    w => w.AssociatedVisitorGroup.Title, 
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id);
                ctx.Load(ctx.Site, s => s.ServerRelativeUrl);

                var masterCatalog = ctx.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                ctx.Load(masterCatalog, m => m.RootFolder.ServerRelativeUrl);

                var themesCatalog = ctx.Web.GetCatalog((int)ListTemplateType.ThemeCatalog);
                ctx.Load(themesCatalog, t => t.RootFolder.ServerRelativeUrl);

                ctx.ExecuteQueryRetry();

                var currentUser = ctx.Web.EnsureProperty(w => w.CurrentUser);


                var ownerGroupName = ctx.Web.AssociatedOwnerGroup.Title;


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
                var associatedOwnerGroupId = parser.ParseString("{groupid:associatedownergroup}");
                var associatedMemberGroupId = parser.ParseString("{groupid:associatedmembergroup}");
                var associatedVisitorGroupId = parser.ParseString("{groupid:associatedvisitorgroup}");
                var groupId = parser.ParseString($"{{groupid:{ownerGroupName}}}");

                Assert.IsTrue(site1 == $"{ctx.Web.ServerRelativeUrl}/test");
                Assert.IsTrue(site2 == $"{ctx.Web.ServerRelativeUrl}/test");
                Assert.IsTrue(sitecol1 == $"{ctx.Site.ServerRelativeUrl}/test");
                Assert.IsTrue(sitecol2 == $"{ctx.Site.ServerRelativeUrl}/test");
                Assert.IsTrue(masterUrl1 == $"{masterCatalog.RootFolder.ServerRelativeUrl}/test");
                Assert.IsTrue(masterUrl2 == $"{masterCatalog.RootFolder.ServerRelativeUrl}/test");
                Assert.IsTrue(themeUrl1 == $"{themesCatalog.RootFolder.ServerRelativeUrl}/test");
                Assert.IsTrue(themeUrl2 == $"{themesCatalog.RootFolder.ServerRelativeUrl}/test");
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
                Assert.IsTrue(int.Parse(associatedOwnerGroupId) == ctx.Web.AssociatedOwnerGroup.Id);
                Assert.IsTrue(int.Parse(associatedMemberGroupId) == ctx.Web.AssociatedMemberGroup.Id);
                Assert.IsTrue(int.Parse(associatedVisitorGroupId) == ctx.Web.AssociatedVisitorGroup.Id);
                Assert.IsTrue(associatedOwnerGroupId == groupId);

            }
        }

        [TestMethod]
        public void RegexSpecialCharactersTests()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.Id, w => w.ServerRelativeUrl, w => w.Title, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                ctx.Load(ctx.Site, s => s.ServerRelativeUrl);

                ctx.ExecuteQueryRetry();

                var currentUser = ctx.Web.EnsureProperty(w => w.CurrentUser);

                ProvisioningTemplate template = new ProvisioningTemplate();
                template.Parameters.Add("test(T)", "test");

                var parser = new TokenParser(ctx.Web, template);

                var web = ctx.Web;

                var contentTypeName = "Test CT (TC) [TC].$";
                var contentTypeId = "0x010801006439AECCDEAE4db2A422A3A04C79CC83";
                var listGuid = Guid.NewGuid();
                var listTitle = @"List (\,*+?|{[()^$.#";
                var listUrl = "Lists/TestList";
                var webPartTitle = @"Webpart (\*+?|{[()^$.#";
                var webPartId = Guid.NewGuid();
                var termSetName = @"Test TermSet (\*+?{[()^$.#";
                var termGroupName = @"Group Name (\*+?{[()^$.#";
                var termStoreName = @"Test TermStore (\*+?{[()^$.#";
                var termSetId = Guid.NewGuid();
                var termStoreId = Guid.NewGuid();


                // Use fake data
                parser.AddToken(new ContentTypeIdToken(web, contentTypeName, contentTypeId));
                parser.AddToken(new ListIdToken(web, listTitle, listGuid));
                parser.AddToken(new ListUrlToken(web, listTitle, listUrl));
                parser.AddToken(new WebPartIdToken(web, webPartTitle, webPartId));
                parser.AddToken(new TermSetIdToken(web, termGroupName, termSetName, termSetId));
                parser.AddToken(new TermStoreIdToken(web, termStoreName, termStoreId));

                var resolvedContentTypeId = parser.ParseString($"{{contenttypeid:{contentTypeName}}}");
                var resolvedListId = parser.ParseString($"{{listid:{listTitle}}}");
                var resolvedListUrl = parser.ParseString($"{{listurl:{listTitle}}}");

                var parameterExpectedResult = $"abc{"test"}/test";
                var parameterTest1 = parser.ParseString("abc{parameter:TEST(T)}/test");
                var parameterTest2 = parser.ParseString("abc{$test(T)}/test");
                var resolvedWebpartId = parser.ParseString($"{{webpartid:{webPartTitle}}}");
                var resolvedTermSetId = parser.ParseString($"{{termsetid:{termGroupName}:{termSetName}}}");
                var resolvedTermStoreId = parser.ParseString($"{{termstoreid:{termStoreName}}}");


                Assert.IsTrue(contentTypeId == resolvedContentTypeId);
                Assert.IsTrue(listUrl == resolvedListUrl);
                Guid outGuid;
                Assert.IsTrue(Guid.TryParse(resolvedListId, out outGuid));
                Assert.IsTrue(parameterTest1 == parameterExpectedResult);
                Assert.IsTrue(parameterTest2 == parameterExpectedResult);
                Assert.IsTrue(Guid.TryParse(resolvedWebpartId, out outGuid));
                Assert.IsTrue(Guid.TryParse(resolvedTermSetId, out outGuid));
                Assert.IsTrue(Guid.TryParse(resolvedTermStoreId, out outGuid));

            }
        }
    }
}
