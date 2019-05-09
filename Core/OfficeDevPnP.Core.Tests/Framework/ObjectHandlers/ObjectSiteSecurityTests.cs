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

        private string originalAssociatedOwnerGroupTitle;
        private string originalAssociatedMemberGroupTitle;
        private string originalAssociatedVisitorGroupTitle;

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
                    w => w.AssociatedVisitorGroup.Id,
                    w => w.AssociatedOwnerGroup.Title,
                    w => w.AssociatedMemberGroup.Title,
                    w => w.AssociatedVisitorGroup.Title);
                ctx.ExecuteQuery();

                originalAssociatedOwnerGroupId = web.AssociatedOwnerGroup.ServerObjectIsNull == true ? -1 : web.AssociatedOwnerGroup.Id;
                originalAssociatedMemberGroupId = web.AssociatedMemberGroup.ServerObjectIsNull == true ? -1 : web.AssociatedMemberGroup.Id;
                originalAssociatedVisitorGroupId = web.AssociatedVisitorGroup.ServerObjectIsNull == true ? -1 : web.AssociatedVisitorGroup.Id;

                originalAssociatedOwnerGroupTitle = web.AssociatedOwnerGroup.ServerObjectIsNull == true ? null : web.AssociatedOwnerGroup.Title;
                originalAssociatedMemberGroupTitle = web.AssociatedMemberGroup.ServerObjectIsNull == true ? null : web.AssociatedMemberGroup.Title;
                originalAssociatedVisitorGroupTitle = web.AssociatedVisitorGroup.ServerObjectIsNull == true ? null : web.AssociatedVisitorGroup.Title;
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

                bool webIsDirty = false;

                if (originalAssociatedOwnerGroupId > 0)
                {
                    Group associatedOwnerGroup = web.SiteGroups.GetById(originalAssociatedOwnerGroupId);
                    associatedOwnerGroup.Title = originalAssociatedOwnerGroupTitle;
                    associatedOwnerGroup.Update();
                    web.AssociatedOwnerGroup = associatedOwnerGroup;
                    webIsDirty = true;
                }

                if (originalAssociatedMemberGroupId > 0)
                {
                    Group associatedMemberGroup = web.SiteGroups.GetById(originalAssociatedMemberGroupId);
                    associatedMemberGroup.Title = originalAssociatedMemberGroupTitle;
                    associatedMemberGroup.Update();
                    web.AssociatedMemberGroup = associatedMemberGroup;
                    webIsDirty = true;
                }

                if (originalAssociatedVisitorGroupId > 0)
                {
                    Group associatedVisitorGroup = web.SiteGroups.GetById(originalAssociatedVisitorGroupId);
                    associatedVisitorGroup.Title = originalAssociatedVisitorGroupTitle;
                    associatedVisitorGroup.Update();
                    web.AssociatedVisitorGroup = associatedVisitorGroup;
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
                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.Security.AdditionalAdministrators.Any());
            }
        }

        [TestMethod]
        public void CanCreateEntities2()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                Web web = ctx.Web;

                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(web) { BaseTemplate = web.GetBaseTemplate() };
                creationInfo.IncludeSiteGroups = true;
                creationInfo.IncludeAssociatedRoleGroups = true;
                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(web, template, creationInfo);

                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Title,
                    w => w.AssociatedMemberGroup.Title,
                    w => w.AssociatedVisitorGroup.Title,
                    w => w.AllProperties);
                ctx.ExecuteQuery();

                Assert.IsTrue(template.Security.AdditionalAdministrators.Any());
                Assert.IsTrue(template.Security.SiteGroups.Any());

#if SP2013
                if (web.AllProperties.FieldValues.ContainsKey("vti_createdassociategroups"))
                {
#else
                    // These three assertions will fail if the site collection does not have the
                    // default groups created during site creation assigned as associated owner group,
                    // associated member group, and associated visitor group.
                    // This is a prerequisite for the site collection used for unit testing purposes.                
                    Assert.IsNull(template.Security.AssociatedOwnerGroup, "Associated owner group is not the default created associated owner group.");
                    Assert.IsNull(template.Security.AssociatedMemberGroup, "Associated owner group is not the default created associated member group.");
                    Assert.IsNull(template.Security.AssociatedVisitorGroup, "Associated owner group is not the default created associated visitor group.");
#endif
#if SP2013
                }
                else
                {
                    Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedOwnerGroup.Title, web), template.Security.AssociatedOwnerGroup, "Associated owner group title mismatch.");
                    Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedMemberGroup.Title, web), template.Security.AssociatedMemberGroup, "Associated member group title mismatch.");
                    Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedVisitorGroup.Title, web), template.Security.AssociatedVisitorGroup, "Associated visitor group title mismatch.");
                }
#endif
            }
        }

        [TestMethod]
        public void CanCreateEntities3()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                InitializeAssociatedGroups(ctx);

                Web web = ctx.Web;

                web.AssociatedOwnerGroup = web.SiteGroups.GetById(ownerGroupId);
                web.AssociatedMemberGroup = web.SiteGroups.GetById(memberGroupId);
                web.AssociatedVisitorGroup = web.SiteGroups.GetById(visitorGroupId);
                web.Update();

                ctx.ExecuteQueryRetry();

                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(web) { BaseTemplate = web.GetBaseTemplate() };
                creationInfo.IncludeSiteGroups = true;
                creationInfo.IncludeAssociatedRoleGroups = true;
                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(web, template, creationInfo);

                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Title,
                    w => w.AssociatedMemberGroup.Title,
                    w => w.AssociatedVisitorGroup.Title);
                ctx.ExecuteQuery();

                Assert.IsTrue(template.Security.AdditionalAdministrators.Any());
                Assert.IsTrue(template.Security.SiteGroups.Any());

                Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedOwnerGroup.Title, web), template.Security.AssociatedOwnerGroup, "Associated owner group title mismatch.");
                Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedMemberGroup.Title, web), template.Security.AssociatedMemberGroup, "Associated member group title mismatch.");
                Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedVisitorGroup.Title, web), template.Security.AssociatedVisitorGroup, "Associated visitor group title mismatch.");
            }
        }

        [TestMethod]
        public void CanSkipExtractSiteGroups()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                Web web = clientContext.Web;
                var creationInfo = new ProvisioningTemplateCreationInformation(web) { BaseTemplate = web.GetBaseTemplate() };
                creationInfo.IncludeSiteGroups = false;
                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(web, template, creationInfo);

                Assert.IsFalse(template.Security.SiteGroups.Any());
            }
        }

        public void CanSkipExtractAssociatedGroups()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                Web web = clientContext.Web;
                var creationInfo = new ProvisioningTemplateCreationInformation(web) { BaseTemplate = web.GetBaseTemplate() };
                creationInfo.IncludeAssociatedRoleGroups = false;
                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(web, template, creationInfo);

                Assert.IsNull(template.Security.AssociatedOwnerGroup);
                Assert.IsNull(template.Security.AssociatedMemberGroup);
                Assert.IsNull(template.Security.AssociatedVisitorGroup);
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

                AssociatedGroupToken associatedOwnerGroupToken = associatedGroupTokens.FirstOrDefault(t => t.GroupType == AssociatedGroupType.Owners);
                AssociatedGroupToken associatedMemberGroupToken = associatedGroupTokens.FirstOrDefault(t => t.GroupType == AssociatedGroupType.Members);
                AssociatedGroupToken associatedVisitorGroupToken = associatedGroupTokens.FirstOrDefault(t => t.GroupType == AssociatedGroupType.Visitors);

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
        public void CanSkipAssigningAssociatedGroups()
        {
            ProvisioningTemplate template = new ProvisioningTemplate();
            template.Security.AssociatedOwnerGroup = "{associatedownergroup}";
            template.Security.AssociatedMemberGroup = "{associatedmembergroup}";
            template.Security.AssociatedVisitorGroup = "{associatedvisitorgroup}";
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
                    w => w.AssociatedVisitorGroup.Id,
                    w => w.AssociatedOwnerGroup.Title,
                    w => w.AssociatedMemberGroup.Title,
                    w => w.AssociatedVisitorGroup.Title);
                ctx.ExecuteQuery();

                Assert.AreEqual(originalAssociatedOwnerGroupId, web.AssociatedOwnerGroup.Id, "Associated owner group ID mismatch.");
                Assert.AreEqual(originalAssociatedMemberGroupId, web.AssociatedMemberGroup.Id, "Associated member group ID mismatch.");
                Assert.AreEqual(originalAssociatedVisitorGroupId, web.AssociatedVisitorGroup.Id, "Associated visitor group ID mismatch.");

                Assert.AreEqual(originalAssociatedOwnerGroupTitle, web.AssociatedOwnerGroup.Title, "Associated owner group Title mismatch.");
                Assert.AreEqual(originalAssociatedMemberGroupTitle, web.AssociatedMemberGroup.Title, "Associated member group Title mismatch.");
                Assert.AreEqual(originalAssociatedVisitorGroupTitle, web.AssociatedVisitorGroup.Title, "Associated visitor group Title mismatch.");
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
                    w => w.AssociatedMemberGroup.Title//,
                    //w => w.AssociatedVisitorGroup
                    );
                clientContext.ExecuteQuery();

                Assert.AreEqual(ownersGroupTitle, web.AssociatedOwnerGroup.Title, "Associated owner group ID mismatch.");
                Assert.AreEqual(membersGroup.Title, web.AssociatedMemberGroup.Title, "Associated member group ID mismatch.");
                Assert.IsTrue(web.AssociatedVisitorGroup.ServerObjectIsNull());
            }
        }
    }
}
