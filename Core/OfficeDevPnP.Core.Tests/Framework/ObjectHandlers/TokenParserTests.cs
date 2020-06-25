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
                ctx.Load(ctx.Site, s => s.ServerRelativeUrl, s => s.Owner);

                var masterCatalog = ctx.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                ctx.Load(masterCatalog, m => m.RootFolder.ServerRelativeUrl);

                var themesCatalog = ctx.Web.GetCatalog((int)ListTemplateType.ThemeCatalog);
                ctx.Load(themesCatalog, t => t.RootFolder.ServerRelativeUrl);

                var expectedRoleDefinitionId = 1073741826;
                var roleDefinition = ctx.Web.RoleDefinitions.GetById(expectedRoleDefinitionId);
                ctx.Load(roleDefinition);

#if SP2019
                if (TestCommon.AppOnlyTesting())
                { 
                    ctx.Web.CreateDefaultAssociatedGroups(ctx.Site.Owner.LoginName, ctx.Site.Owner.LoginName, string.Empty); 
                }
                else
                {
                    ctx.Web.CreateDefaultAssociatedGroups(string.Empty, string.Empty, string.Empty);
                } 
#endif

                ctx.ExecuteQueryRetry();

                var currentUser = ctx.Web.EnsureProperty(w => w.CurrentUser);


                var ownerGroupName = ctx.Web.AssociatedOwnerGroup.Title;


                ProvisioningTemplate template = new ProvisioningTemplate();
                template.Parameters.Add("test", "test");
                template.Parameters.Add("test2", "test2");
                // Due to the refactoring of the parser only tokens specified in the template are loaded
                template.Parameters.Add("sitename", "{sitename}");
                template.Parameters.Add("siteid", "{siteid}");
                template.Parameters.Add("site", "{site}");
                template.Parameters.Add("sitecollection", "{sitecollection}");
                template.Parameters.Add("masterpagecatalog", "{masterpagecatalog}");
                template.Parameters.Add("themecatalog", "{themecatalog}");
                template.Parameters.Add("associatedownergroup", "{associatedownergroup}");
                template.Parameters.Add("associatedmembergroup", "{associatedmembergroup}");
                template.Parameters.Add("associatedvisitorgroup", "{associatedvisitorgroup}");
                template.Parameters.Add("currentuserid", "{currentuserid}");
                template.Parameters.Add("currentuserloginname", "{currentuserloginname}");
                template.Parameters.Add("currentuserfullname", "{currentuserfullname}");
                template.Parameters.Add("guid", "{guid}");
                template.Parameters.Add("groupid:associatedownergroup", "{groupid:associatedownergroup}");
                template.Parameters.Add("associatedownergroupid", "{associatedownergroupid}");
                template.Parameters.Add("siteowner", "{siteowner}");
                template.Parameters.Add("everyonebutexternalusers", "{everyonebutexternalusers}");
                template.Parameters.Add("roledefinitionid", "{roledefinitionid}");
                template.Parameters.Add("termid", "{termsetid:{parameter:test}:{parameter:test}}");

                var parser = new TokenParser(ctx.Web, template);
                parser.AddToken(new FieldIdToken(ctx.Web, "DemoField", new Guid("7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0")));
              
                var siteName = parser.ParseString("{sitename}");
                var siteId = parser.ParseString("{siteid}");
                var site = parser.ParseString("{site}/test");
                var sitecol = parser.ParseString("{sitecollection}/test");
                var masterUrl = parser.ParseString("{masterpagecatalog}/test");
                var themeUrl = parser.ParseString("{themecatalog}/test");
                var parameterTest = parser.ParseString("abc{parameter:TEST}/test");
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
                var siteOwner = parser.ParseString("{siteowner}");
                var roleDefinitionId = parser.ParseString($"{{roledefinitionid:{roleDefinition.Name}}}");
                var xamlEscapeString = "{}{0}Id";
                var parsedXamlEscapeString = parser.ParseString(xamlEscapeString);
                const string fieldRef = @"<FieldRefs><FieldRef Name=""DemoField"" ID=""{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}"" /></FieldRefs></Field>";
                var parsedFieldRef = parser.ParseString(@"<FieldRefs><FieldRef Name=""DemoField"" ID=""{{fieldid:DemoField}}"" /></FieldRefs></Field>");
                var everyoneExceptExternals = parser.ParseString("{everyonebutexternalusers}");
                
                

                Assert.IsTrue(site == $"{ctx.Web.ServerRelativeUrl}/test");
                Assert.IsTrue(sitecol == $"{ctx.Site.ServerRelativeUrl}/test");
                Assert.IsTrue(masterUrl == $"{masterCatalog.RootFolder.ServerRelativeUrl}/test");
                Assert.IsTrue(themeUrl == $"{themesCatalog.RootFolder.ServerRelativeUrl}/test");
                Assert.IsTrue(parameterTest == "abctest/test");
                Assert.IsTrue(associatedOwnerGroup == ctx.Web.AssociatedOwnerGroup.Title);
                Assert.IsTrue(associatedVisitorGroup == ctx.Web.AssociatedVisitorGroup.Title);
                Assert.IsTrue(associatedMemberGroup == ctx.Web.AssociatedMemberGroup.Title);
                Assert.IsTrue(siteName == ctx.Web.Title);
                Assert.IsTrue(siteId == ctx.Web.Id.ToString());
                Assert.IsTrue(currentUserId == currentUser.Id.ToString());
                Assert.IsTrue(currentUserFullName == currentUser.Title);
                Assert.IsTrue(currentUserLoginName.Equals(currentUser.LoginName, StringComparison.OrdinalIgnoreCase));
                // Guid token is 
                Guid outGuid;
                Assert.IsTrue(Guid.TryParse(guid, out outGuid));
                Assert.IsTrue(int.Parse(associatedOwnerGroupId) == ctx.Web.AssociatedOwnerGroup.Id);
                Assert.IsTrue(int.Parse(associatedMemberGroupId) == ctx.Web.AssociatedMemberGroup.Id);
                Assert.IsTrue(int.Parse(associatedVisitorGroupId) == ctx.Web.AssociatedVisitorGroup.Id);
                Assert.IsTrue(associatedOwnerGroupId == groupId);
                Assert.IsTrue(siteOwner == ctx.Site.Owner.LoginName);
                Assert.IsTrue(roleDefinitionId == expectedRoleDefinitionId.ToString(), $"Role Definition Id was not parsed correctly (expected:{expectedRoleDefinitionId};returned:{roleDefinitionId})");
                Assert.IsTrue(parsedXamlEscapeString == xamlEscapeString);
                Assert.IsTrue(parsedFieldRef.ToUpperInvariant() == fieldRef.ToUpperInvariant());
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
                template.Parameters.Add("a{b", "test");

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

                var resolvedTermGroupName = parser.ParseString("{sitecollectiontermgroupname}");

                // Use fake data
                parser.AddToken(new ContentTypeIdToken(web, contentTypeName, contentTypeId));
                parser.AddToken(new ListIdToken(web, listTitle, listGuid));
                parser.AddToken(new ListUrlToken(web, listTitle, listUrl));
                parser.AddToken(new WebPartIdToken(web, webPartTitle, webPartId));
                parser.AddToken(new TermSetIdToken(web, termGroupName, termSetName, termSetId));
                parser.AddToken(new TermSetIdToken(web, resolvedTermGroupName, termSetName, termSetId));
                parser.AddToken(new TermStoreIdToken(web, termStoreName, termStoreId));

                var resolvedContentTypeId = parser.ParseString($"{{contenttypeid:{contentTypeName}}}");
                var resolvedListId = parser.ParseString($"{{listid:{listTitle}}}");
                var resolvedListUrl = parser.ParseString($"{{listurl:{listTitle}}}");

                var parameterExpectedResult = $"abc{"test"}/test";
                var parameterTest1 = parser.ParseString("abc{parameter:TEST(T)}/test");
                var parameterTest2 = parser.ParseString("abc{parameter:a{b}/test");
                var resolvedWebpartId = parser.ParseString($"{{webpartid:{webPartTitle}}}");
                var resolvedTermSetId = parser.ParseString($"{{termsetid:{termGroupName}:{termSetName}}}");
                var resolvedTermSetId2 = parser.ParseString($"{{termsetid:{{sitecollectiontermgroupname}}:{termSetName}}}");
                var resolvedTermStoreId = parser.ParseString($"{{termstoreid:{termStoreName}}}");


                Assert.IsTrue(contentTypeId == resolvedContentTypeId);
                Assert.IsTrue(listUrl == resolvedListUrl);
                Guid outGuid;
                Assert.IsTrue(Guid.TryParse(resolvedListId, out outGuid));
                Assert.IsTrue(parameterTest1 == parameterExpectedResult);
                Assert.IsTrue(parameterTest2 == parameterExpectedResult);
                Assert.IsTrue(Guid.TryParse(resolvedWebpartId, out outGuid));
                Assert.IsTrue(Guid.TryParse(resolvedTermSetId, out outGuid));
                Assert.IsTrue(Guid.TryParse(resolvedTermSetId2, out outGuid));
                Assert.IsTrue(Guid.TryParse(resolvedTermStoreId, out outGuid));

            }
        }

        [TestMethod]
        [Timeout(1 * 60 * 1000)]
        public void NestedTokenTests()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.Id, w => w.ServerRelativeUrl, w => w.Title, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                ctx.Load(ctx.Site, s => s.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();

                var template = new ProvisioningTemplate();
                template.Parameters.Add("test3", "reallynestedvalue:{parameter:test2}");
                template.Parameters.Add("test1", "testValue");
                template.Parameters.Add("test2", "value:{parameter:test1}");

                // Test parameters that infinitely loop
                template.Parameters.Add("chain1", "{parameter:chain3}");
                template.Parameters.Add("chain2", "{parameter:chain1}");
                template.Parameters.Add("chain3", "{parameter:chain2}");

                var parser = new TokenParser(ctx.Web, template);

                var parameterTest1 = parser.ParseString("parameterTest:{parameter:test1}");
                var parameterTest2 = parser.ParseString("parameterTest:{parameter:test2}");
                var parameterTest3 = parser.ParseString("parameterTest:{parameter:test3}");
                Assert.IsTrue(parameterTest1 == "parameterTest:testValue");
                Assert.IsTrue(parameterTest2 == "parameterTest:value:testValue");
                Assert.IsTrue(parameterTest3 == "parameterTest:reallynestedvalue:value:testValue");

                var chainTest1 = parser.ParseString("parameterTest:{parameter:chain1}");

                // Parser should stop processing parent tokens when processing nested tokens,
                // so we should end up with the value of the last param (chain2) in our param chain,
                // which will not get detokenized.
                Assert.IsTrue(chainTest1 == "parameterTest:{parameter:chain1}");
            }
        }

        [TestMethod]
        public void WebPartTests()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.Id, w => w.ServerRelativeUrl, w => w.Title, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                ctx.Load(ctx.Site, s => s.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();

                var listGuid = Guid.NewGuid();
                var listTitle = "MyList";
                var web = ctx.Web;

                ProvisioningTemplate template = new ProvisioningTemplate();
                template.Parameters.Add("test", "test");
                var parser = new TokenParser(ctx.Web, template);
                parser.AddToken(new ListIdToken(web, listTitle, listGuid));
                parser.ParseString($"{{listid:{listTitle}}}");
                var listId = parser.ParseStringWebPart($"{{listid:{listTitle}}}", web, null);
                Assert.IsTrue(listGuid.ToString() == listId);

                var parameterValue = parser.ParseStringWebPart("{parameter:test}", web, null);
                Assert.IsTrue("test" == parameterValue);
            }
        }
    }
}