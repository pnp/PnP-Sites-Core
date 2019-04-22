using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;
using User = OfficeDevPnP.Core.Framework.Provisioning.Model.User;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectSiteSecurityTests
    {
        private List<UserEntity> admins;
        private readonly string ownerGroupName;
        private readonly string memberGroupName;
        private readonly string visitorGroupName;
        private readonly string testWebName;

        private int ownerGroupId;
        private int memberGroupId;
        private int visitorGroupId;

        private int originalAssociatedOwnerGroupId;
        private int originalAssociatedMemberGroupId;
        private int originalAssociatedVisitorGroupId;

        public ObjectSiteSecurityTests()
        {
            ownerGroupName = string.Format("Test_Owner Group_{0}", DateTime.Now.Ticks);
            memberGroupName = string.Format("Test_Member Group_{0}", DateTime.Now.Ticks);
            visitorGroupName = string.Format("Test_Visitor Group_{0}", DateTime.Now.Ticks);
            testWebName = string.Format("Test_{0:yyyyMMddTHHmmss}", DateTimeOffset.Now);
        }

        [TestInitialize]
        public void Initialize()
        {

            using (var ctx = TestCommon.CreateClientContext())
            {
                admins = ctx.Web.GetAdministrators();

                Web web = ctx.Web;

                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id);
                ctx.ExecuteQuery();

                originalAssociatedOwnerGroupId = web.AssociatedOwnerGroup.ServerObjectIsNull == true ? -1 : web.AssociatedOwnerGroup.Id;
                originalAssociatedMemberGroupId = web.AssociatedMemberGroup.ServerObjectIsNull == true ? -1 : web.AssociatedMemberGroup.Id;
                originalAssociatedVisitorGroupId = web.AssociatedVisitorGroup.ServerObjectIsNull == true ? -1 : web.AssociatedVisitorGroup.Id;
            }
        }

        public void InitializeAssociatedGroups(ClientContext ctx)
        {
            Web web = ctx.Web;

            Group ownerGroup = web.SiteGroups.Add(new GroupCreationInformation
            {
                Title = ownerGroupName
            });

            Group memberGroup = web.SiteGroups.Add(new GroupCreationInformation
            {
                Title = memberGroupName
            });

            Group visitorGroup = web.SiteGroups.Add(new GroupCreationInformation
            {
                Title = visitorGroupName
            });

            ctx.Load(ownerGroup, og => og.Id);
            ctx.Load(memberGroup, mg => mg.Id);
            ctx.Load(visitorGroup, vg => vg.Id);
            ctx.ExecuteQuery();

            ownerGroupId = ownerGroup.Id;
            memberGroupId = memberGroup.Id;
            visitorGroupId = visitorGroup.Id;
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var memberGroup = ctx.Web.AssociatedMemberGroup;
                ctx.Load(memberGroup);
                ctx.ExecuteQueryRetry();
                if (memberGroup.ServerObjectIsNull == false)
                {
                    foreach (var user in admins)
                    {
                        try
                        {
                            ctx.Web.RemoveUserFromGroup(memberGroup.Title, user.LoginName);
                        }
                        catch (ServerException)
                        {

                        }
                    }
                }

                var siteGroups = ctx.Web.SiteGroups;
                ctx.Load(siteGroups, sg => sg.Include(g => g.Title));
                ctx.ExecuteQueryRetry();
                IList<Group> groupsToRemove = new List<Group>();

                foreach (var group in siteGroups)
                {
                    if (group.Title.StartsWith("Test_"))
                    {
                        groupsToRemove.Add(group);
                    }
                }
                if (groupsToRemove.Count > 0)
                {
                    foreach (var group in groupsToRemove)
                    {
                        ctx.Web.SiteGroups.Remove(group);
                    }
                    ctx.ExecuteQueryRetry();
                }
                Web web = ctx.Web;

                //ConditionalScope associatedOwnerScope = new ConditionalScope(ctx, () => web.AssociatedOwnerGroup.ServerObjectIsNull == false);
                //using (associatedOwnerScope.StartScope())
                //{
                //    ctx.Load(web, w => w.AssociatedOwnerGroup.Id);
                //}

                //ConditionalScope associatedMemberScope = new ConditionalScope(ctx, () => web.AssociatedMemberGroup.ServerObjectIsNull == false);
                //using (associatedMemberScope.StartScope())
                //{
                //    ctx.Load(web, w => w.AssociatedMemberGroup.Id);
                //}

                //ConditionalScope associatedVisitorScope = new ConditionalScope(ctx, () => web.AssociatedVisitorGroup.ServerObjectIsNull == false);
                //using (associatedVisitorScope.StartScope())
                //{
                //    ctx.Load(web, w => w.AssociatedVisitorGroup.Id);
                //}

                //ctx.ExecuteQuery();

                bool webIsDirty = false;

                if (originalAssociatedOwnerGroupId > 0)
                {
                    web.AssociatedOwnerGroup = web.SiteGroups.GetById(originalAssociatedOwnerGroupId);
                    webIsDirty = true;
                }

                if (originalAssociatedMemberGroupId > 0)
                {
                    web.AssociatedMemberGroup = web.SiteGroups.GetById(originalAssociatedMemberGroupId);
                    webIsDirty = true;
                }

                if (originalAssociatedVisitorGroupId > 0)
                {
                    web.AssociatedVisitorGroup = web.SiteGroups.GetById(originalAssociatedVisitorGroupId);
                    webIsDirty = true;
                }

                if (webIsDirty)
                {
                    web.Update();
                }

                var subWebs = ctx.Web.Webs;
                ctx.Load(subWebs, wc => wc.Include(w => w.Title, w => w.ServerRelativeUrl));
                ctx.ExecuteQueryRetry();

                var websToDelete = new List<Web>();

                foreach (var subWeb in subWebs)
                {
                    if (subWeb.Title.StartsWith("Test_"))
                    {
                        websToDelete.Add(subWeb);
                    }
                }

                foreach (var webToDelete in websToDelete)
                {
                    Console.WriteLine("Deleting site {0}", webToDelete.ServerRelativeUrl);
                    webToDelete.DeleteObject();
                    try
                    {
                        ctx.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Exception cleaning up: {0}", ex);
                    }
                }
            }
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();
            var roleDefinitionName = "UT_RoleDefinition";

            foreach (var user in admins)
            {
                template.Security.AdditionalMembers.Add(new User() { Name = user.LoginName });
            }
            template.Security.SiteSecurityPermissions.RoleDefinitions.Add(new Core.Framework.Provisioning.Model.RoleDefinition()
            {
                Name = roleDefinitionName
            });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var memberGroup = ctx.Web.AssociatedMemberGroup;
                var roleDefinitions = ctx.Web.RoleDefinitions;
                ctx.Load(memberGroup, g => g.Users);
                ctx.Load(roleDefinitions);
                ctx.ExecuteQueryRetry();
                foreach (var user in admins)
                {
                    var existingUser = memberGroup.Users.GetByLoginName(user.LoginName);
                    ctx.Load(existingUser);
                    ctx.ExecuteQueryRetry();
                    Assert.IsNotNull(existingUser);
                }
                Assert.IsTrue(roleDefinitions.Any(rd => rd.Name == roleDefinitionName),"New role definition wasn't found after provisioning");
            }
        }

        [TestMethod]
        public void CanCreateEntities1()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                Web web = ctx.Web;
                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Title,
                    w => w.AssociatedMemberGroup.Title,
                    w => w.AssociatedVisitorGroup.Title);
                ctx.ExecuteQuery();

                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.Security.AdditionalAdministrators.Any());
                Assert.AreEqual(web.AssociatedOwnerGroup.Title, template.Security.AssociatedOwnerGroup);
                Assert.AreEqual(web.AssociatedMemberGroup.Title, template.Security.AssociatedMemberGroup);
                Assert.AreEqual(web.AssociatedVisitorGroup.Title, template.Security.AssociatedVisitorGroup);
            }
        }

        [TestMethod]
        public void CanCreateEntities2()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };
                creationInfo.IncludeSiteGroups = true;
                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.Security.AdditionalAdministrators.Any());
                Assert.IsTrue(template.Security.SiteGroups.Any());

            }
        }

        [TestMethod]
        public void CanProvisionSiteGroupDescription()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var currentUser = ctx.Web.EnsureProperty(w => w.CurrentUser);
                var currentUserLoginName = currentUser.EnsureProperty(u => u.LoginName);

                string plainTextGroupName = string.Format("Test_PlainText_{0}", DateTime.Now.Ticks);
                string htmlGroupName = string.Format("Test_HTML_{0}", DateTime.Now.Ticks);
                string plainTextDescription = "Testing plain text";
                string richHtmlDescription = "Testing HTML with link to <a href=\"/\">Root Site</a>";
                string richHtmlDescriptionPlainTextVersion = "Testing HTML with link to Root Site";

                var template = new ProvisioningTemplate();

                template.Security.SiteGroups.Add(new SiteGroup()
                {
                    Title = plainTextGroupName,
                    Description = plainTextDescription,
                    Owner = currentUserLoginName,
                });

                template.Security.SiteGroups.Add(new SiteGroup()
                {
                    Title = htmlGroupName,
                    Description = richHtmlDescription,
                    Owner = currentUserLoginName,
                });

                var parser = new TokenParser(ctx.Web, template);
                new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var plainTextGroup = ctx.Web.SiteGroups.GetByName(plainTextGroupName);
                var htmlGroup = ctx.Web.SiteGroups.GetByName(htmlGroupName);
                ctx.Load(plainTextGroup, g => g.Id, g => g.Description);
                ctx.Load(htmlGroup, g => g.Id, g => g.Description);
                ctx.ExecuteQueryRetry();

                var plainTextGroupItem = ctx.Web.SiteUserInfoList.GetItemById(plainTextGroup.Id);
                var htmlGroupItem = ctx.Web.SiteUserInfoList.GetItemById(htmlGroup.Id);
                ctx.Web.Context.Load(plainTextGroupItem, g => g["Notes"]);
                ctx.Web.Context.Load(htmlGroupItem, g => g["Notes"]);
                ctx.Web.Context.ExecuteQueryRetry();

                Assert.AreEqual(plainTextDescription, plainTextGroup.Description);
                Assert.AreEqual(plainTextDescription, plainTextGroupItem["Notes"]);

                Assert.AreEqual(richHtmlDescriptionPlainTextVersion, htmlGroup.Description);
                Assert.AreEqual(richHtmlDescription, htmlGroupItem["Notes"]);
            }
        }

        [TestMethod()]
        public void CanProvisionAssociatedGroups()
        {
            ProvisioningTemplate template = new ProvisioningTemplate();
            template.Security.AssociatedOwnerGroup = ownerGroupName;
            template.Security.AssociatedMemberGroup = memberGroupName;
            template.Security.AssociatedVisitorGroup = visitorGroupName;
            foreach (var user in admins)
            {
                template.Security.AdditionalMembers.Add(new User() { Name = user.LoginName });
            }

            using (var ctx = TestCommon.CreateClientContext())
            {
                InitializeAssociatedGroups(ctx);
                Web web = ctx.Web;

                var parser = new TokenParser(ctx.Web, template);
                new ObjectSiteSecurity().ProvisionObjects(web, template, parser, new ProvisioningTemplateApplyingInformation());

                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id);
                ctx.ExecuteQuery();

                Assert.AreEqual(ownerGroupId, web.AssociatedOwnerGroup.Id, "Associated owner group ID mismatch.");
                Assert.AreEqual(memberGroupId, web.AssociatedMemberGroup.Id, "Associated member group ID mismatch.");
                Assert.AreEqual(visitorGroupId, web.AssociatedVisitorGroup.Id, "Associated visitor group ID mismatch.");
                IEnumerable<AssociatedGroupToken> associatedGroupTokens = parser.Tokens.Where(t => t.GetType() == typeof(AssociatedGroupToken)).Cast<AssociatedGroupToken>();

                AssociatedGroupToken associatedOwnerGroupToken = associatedGroupTokens.FirstOrDefault(t => t.GroupType == AssociatedGroupToken.AssociatedGroupType.owners);
                AssociatedGroupToken associatedMemberGroupToken = associatedGroupTokens.FirstOrDefault(t => t.GroupType == AssociatedGroupToken.AssociatedGroupType.members);
                AssociatedGroupToken associatedVisitorGroupToken = associatedGroupTokens.FirstOrDefault(t => t.GroupType == AssociatedGroupToken.AssociatedGroupType.visitors);

                Assert.IsNotNull(associatedOwnerGroupToken);
                Assert.IsNotNull(associatedMemberGroupToken);
                Assert.IsNotNull(associatedVisitorGroupToken);

                Assert.AreEqual(ownerGroupName, associatedOwnerGroupToken.GetReplaceValue());
                Assert.AreEqual(memberGroupName, associatedMemberGroupToken.GetReplaceValue());
                Assert.AreEqual(visitorGroupName, associatedVisitorGroupToken.GetReplaceValue());

                foreach (var user in admins)
                {
                    var existingUser = web.AssociatedMemberGroup.Users.GetByLoginName(user.LoginName);
                    ctx.Load(existingUser);
                    ctx.ExecuteQueryRetry();
                    Assert.IsNotNull(existingUser);
                }

            }
        }

        [TestMethod()]
        public void CanProvisionAssociatedGroupsInSubSite()
        {
            ProvisioningTemplate template = new ProvisioningTemplate();
            template.Security.AssociatedOwnerGroup = ownerGroupName;
            template.Security.AssociatedMemberGroup = memberGroupName;
            template.Security.AssociatedVisitorGroup = visitorGroupName;
            template.Security.BreakRoleInheritance = true;
            template.Security.CopyRoleAssignments = false;
            template.Security.ClearSubscopes = true;

            using (var clientContext = TestCommon.CreateClientContext())
            {
                InitializeAssociatedGroups(clientContext);

                Web web = clientContext.Web;
                Web subWeb = web.CreateWeb(testWebName, testWebName, "Test", "STS#0", 1033);
                clientContext.ExecuteQuery();

                var parser = new TokenParser(clientContext.Web, template);
                new ObjectSiteSecurity().ProvisionObjects(subWeb, template, parser, new ProvisioningTemplateApplyingInformation());

                subWeb.Context.Load(subWeb,
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id);
                clientContext.ExecuteQuery();

                Assert.AreEqual(ownerGroupId, subWeb.AssociatedOwnerGroup.Id, "Associated owner group ID mismatch.");
                Assert.AreEqual(memberGroupId, subWeb.AssociatedMemberGroup.Id, "Associated member group ID mismatch.");
                Assert.AreEqual(visitorGroupId, subWeb.AssociatedVisitorGroup.Id, "Associated visitor group ID mismatch.");
            }
        }

        [TestMethod()]
        public void CanProvisionNewAssociatedGroup()
        {
            ProvisioningTemplate template = new ProvisioningTemplate();

            SiteGroup membersGroup = new SiteGroup()
            {
                Title = string.Format("Test_New Group_{0}", DateTime.Now.Ticks),
            };
            string ownersGroupTitle = string.Format("Test_New Group2_{0}", DateTime.Now.Ticks);

            template.Security.SiteGroups.Add(membersGroup);
            template.Security.AssociatedMemberGroup = membersGroup.Title;
            template.Security.AssociatedOwnerGroup = ownersGroupTitle;
            template.Security.AssociatedVisitorGroup = "";

            using (var clientContext = TestCommon.CreateClientContext())
            {
                Web web = clientContext.Web;

                var parser = new TokenParser(clientContext.Web, template);
                new ObjectSiteSecurity().ProvisionObjects(web, template, parser, new ProvisioningTemplateApplyingInformation());
            }
            using (var clientContext = TestCommon.CreateClientContext())
            {
                Web web = clientContext.Web;

                clientContext.Load(web,
                    w => w.AssociatedOwnerGroup.Title,
                    w => w.AssociatedMemberGroup.Title,
                    w => w.AssociatedVisitorGroup);
                clientContext.ExecuteQuery();

                Assert.AreEqual(ownersGroupTitle, web.AssociatedOwnerGroup.Title, "Associated owner group ID mismatch.");
                Assert.AreEqual(membersGroup.Title, web.AssociatedMemberGroup.Title, "Associated member group ID mismatch.");
                Assert.IsTrue(web.AssociatedVisitorGroup.ServerObjectIsNull.Value);
            }
        }
    }
}
