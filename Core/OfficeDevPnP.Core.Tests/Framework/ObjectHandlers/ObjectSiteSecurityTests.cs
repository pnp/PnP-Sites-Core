using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
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

        private bool originalIsNoScriptSite;

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
            TestCommon.FixAssemblyResolving("Newtonsoft.Json");

            using (var ctx = TestCommon.CreateClientContext())
            {
                admins = ctx.Web.GetAdministrators();

                Web web = ctx.Web;

                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id,
                    w => w.Url);
                ctx.ExecuteQueryRetry();

                originalAssociatedOwnerGroupId = web.AssociatedOwnerGroup.ServerObjectIsNull == true ? -1 : web.AssociatedOwnerGroup.Id;
                originalAssociatedMemberGroupId = web.AssociatedMemberGroup.ServerObjectIsNull == true ? -1 : web.AssociatedMemberGroup.Id;
                originalAssociatedVisitorGroupId = web.AssociatedVisitorGroup.ServerObjectIsNull == true ? -1 : web.AssociatedVisitorGroup.Id;
                originalIsNoScriptSite = web.IsNoScriptSite();
#if !SP2013 && !SP2016
                if (originalIsNoScriptSite)
                {
                    AllowScripting(web.Url, true);
                }
#endif
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
            ctx.ExecuteQueryRetry();

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
                ctx.Load(ctx.Web, w => w.Url);
                ctx.ExecuteQueryRetry();
#if !SP2013 && !SP2016
                AllowScripting(ctx.Web.Url, !originalIsNoScriptSite);
#endif
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

                var websToDelete = new List<Web>();
                try { 
                    var subWebs = ctx.Web.Webs; // this call fails in NoScript sites
                    ctx.Load(subWebs, wc => wc.Include(w => w.Title, w => w.ServerRelativeUrl));
                    ctx.ExecuteQueryRetry();

                    foreach (var subWeb in subWebs)
                    {
                        if (subWeb.Title.StartsWith("Test_"))
                        {
                            websToDelete.Add(subWeb);
                        }
                    }
                } catch (Exception e)
                {
                    Console.WriteLine("Error while accessing subwebs: " + e.ToDetailedString());
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

            SiteGroup membersGroup = new SiteGroup()
            {
                Title = string.Format("Test_New Group_{0}", DateTime.Now.Ticks),
            };
            template.Security.SiteGroups.Add(membersGroup);
            template.Security.AssociatedMemberGroup = membersGroup.Title;

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
                ctx.Load(ctx.Web, p => p.AssociatedMemberGroup);
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
                var template = new ProvisioningTemplate();
                template = new ObjectSiteSecurity().ExtractObjects(web, template, creationInfo);

                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Title,
                    w => w.AssociatedMemberGroup.Title,
                    w => w.AssociatedVisitorGroup.Title,
                    w => w.Url /* needed by TokenContext somewhere... */);
                ctx.ExecuteQueryRetry();

                Assert.IsTrue(template.Security.AdditionalAdministrators.Any());
                // Assert.IsFalse(template.Security.SiteGroups.Any());
                Assert.IsTrue(template.Security.SiteGroups.Any());

                // tbd: fix those...
                //Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedOwnerGroup.Title, web), template.Security.AssociatedOwnerGroup, "Associated owner group title mismatch.");
                //Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedMemberGroup.Title, web), template.Security.AssociatedMemberGroup, "Associated member group title mismatch.");
                //Assert.AreEqual(SiteTitleToken.GetReplaceToken(web.AssociatedVisitorGroup.Title, web), template.Security.AssociatedVisitorGroup, "Associated visitor group title mismatch.");

                // These three assertions will fail if the site collection does not have the
                // default groups created during site creation assigned as associated owner group,
                // associated member group, and associated visitor group.
                // This is a prerequisite for the site collection used for unit testing purposes.        
                
                // Makes no sense to evaluate this as the site collection can perfectly have no associated groups set due to other tests that ran/failed before
                //Assert.IsTrue(template.Security.AssociatedOwnerGroup.Contains("{groupsitetitle}"), "Associated owner group title does not contain the Group Site Title token.");
                //Assert.IsTrue(template.Security.AssociatedMemberGroup.Contains("{groupsitetitle}"), "Associated owner group title does not contain the Group Site Title token.");
                //Assert.IsTrue(template.Security.AssociatedVisitorGroup.Contains("{groupsitetitle}"), "Associated owner group title does not contain the Group Site Title token.");
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
            template.Security.SiteGroups.Add(new SiteGroup()
            {
                Title = ownerGroupName
            });
            template.Security.SiteGroups.Add(new SiteGroup()
            {
                Title = memberGroupName
            });
            template.Security.SiteGroups.Add(new SiteGroup()
            {
                Title = visitorGroupName
            });
            foreach (var user in admins)
            {
                template.Security.AdditionalMembers.Add(new User() { Name = user.LoginName });
            }

            using (var ctx = TestCommon.CreateClientContext())
            {
                InitializeAssociatedGroups(ctx);
                Web web = ctx.Web;

                var parser = new TokenParser(ctx.Web, template);
                if (parser.Tokens.Count == 0)
                {
                    parser.AddToken(new AssociatedGroupIdToken(web, AssociatedGroupIdToken.AssociatedGroupType.owners));
                    parser.AddToken(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.owners));

                    parser.AddToken(new AssociatedGroupIdToken(web, AssociatedGroupIdToken.AssociatedGroupType.members));
                    parser.AddToken(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.members));

                    parser.AddToken(new AssociatedGroupIdToken(web, AssociatedGroupIdToken.AssociatedGroupType.visitors));
                    parser.AddToken(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.visitors));
                }

                new ObjectSiteSecurity().ProvisionObjects(web, template, parser, new ProvisioningTemplateApplyingInformation());

                ctx.Load(web,
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id);
                ctx.ExecuteQueryRetry();

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
                Title = string.Format("Test_New MemberGroup_{0}", DateTime.Now.Ticks),
            };
            string ownersGroupTitle = string.Format("Test_New OwnerGroup2_{0}", DateTime.Now.Ticks);

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

                if (!web.AssociatedOwnerGroup.ServerObjectIsNull())
                {
                    web.AssociatedOwnerGroup.EnsureProperty(g => g.Title);
                }

                if (!web.AssociatedMemberGroup.ServerObjectIsNull())
                {
                    web.AssociatedMemberGroup.EnsureProperty(g => g.Title);
                }

                if (!web.AssociatedVisitorGroup.ServerObjectIsNull())
                {
                    web.AssociatedVisitorGroup.EnsureProperty(g => g.Title);
                }

                Assert.AreNotEqual(ownersGroupTitle, web.AssociatedOwnerGroup.Title, "Associated owner group ID mismatch.");
                Assert.AreEqual(membersGroup.Title, web.AssociatedMemberGroup.Title, "Associated member group ID mismatch.");
                //Assert.IsTrue(web.AssociatedVisitorGroup.ServerObjectIsNull());
            }
        }

#if !SP2013 && !SP2016
        // ensure #2127 does not occur again; specifically check that not too many groups are created
        [TestMethod()]
        public async Task CanExportAndImportAssociatedGroupsProperlyInNewNoScriptSite()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Site creation requires owner, so this will not yet work in app-only");

            this.CheckAndExcludeTestNotSupportedForAppOnlyTestting();

            var newCommSiteUrl = "";
            var newSiteTitle = "Comm Site Test - Groups";
            var loginName = "";
            ProvisioningTemplate template;
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // get the template from real site as this produced the template that lead to too many groups being created which lead to #2127
                var creationInfo = new ProvisioningTemplateCreationInformation(clientContext.Web);
                creationInfo.HandlersToProcess = Handlers.SiteSecurity;
                template = clientContext.Web.GetProvisioningTemplate();

                var user = clientContext.Web.CurrentUser;
                clientContext.Load(user);
                clientContext.ExecuteQueryRetry();
                loginName = user.LoginName;
                template.Security.AdditionalMembers.Add(new User() { Name = loginName });
                template.Security.AdditionalOwners.Add(new User() { Name = loginName });
                template.Security.AdditionalVisitors.Add(new User() { Name = loginName });

                newCommSiteUrl = await CreateCommunicationSite(clientContext, newSiteTitle);
            }
            try
            {
                using (var ctx = TestCommon.CreateClientContext(newCommSiteUrl))
                {
                    Web web = ctx.Web;

                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectSiteSecurity().ProvisionObjects(web, template, parser, new ProvisioningTemplateApplyingInformation());

#if SP2019
                    await WaitForAsyncGroupTitleChangeWithTimeout(ctx);
#endif

                    ctx.Load(web,
                        w => w.SiteGroups,
                        w => w.AssociatedOwnerGroup.Users,
                        w => w.AssociatedOwnerGroup.Title,
                        w => w.AssociatedMemberGroup.Users,
                        w => w.AssociatedMemberGroup.Title,
                        w => w.AssociatedVisitorGroup.Users,
                        w => w.AssociatedVisitorGroup.Title);
                    ctx.ExecuteQueryRetry();

                    Assert.AreEqual(3, web.SiteGroups.Count, "Unexpected number of groups found");
                    Assert.AreEqual(1, web.AssociatedVisitorGroup.Users.Count(u => u.LoginName == loginName));
                    Assert.AreEqual(1, web.AssociatedMemberGroup.Users.Count(u => u.LoginName == loginName));
                    Assert.AreEqual(1, web.AssociatedOwnerGroup.Users.Count(u => u.LoginName == loginName));
                    Assert.AreEqual(newSiteTitle + " Visitors", web.AssociatedVisitorGroup.Title);
                    Assert.AreEqual(newSiteTitle + " Members", web.AssociatedMemberGroup.Title);
                    Assert.AreEqual(newSiteTitle + " Owners", web.AssociatedOwnerGroup.Title);
                }
            } 
            finally
            {
                using (var clientContext = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(clientContext);
#if !ONPREMISES
                    tenant.DeleteSiteCollection(newCommSiteUrl, false);
#else
                    tenant.DeleteSiteCollection(newCommSiteUrl);
#endif
                }
            }
        }

        // for details see here: https://github.com/SharePoint/PnP-Sites-Core/pull/2174#issuecomment-487538551
        [TestMethod()]
        public async Task CanMapAssociatedGroupsToExistingOnesInNewScriptSite()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Site creation requires owner, so this will not yet work in app-only");

            var newCommSiteUrl = string.Empty;

            ProvisioningTemplate template = new ProvisioningTemplate();
            template.Security.AssociatedOwnerGroup = "These names shouldn't matter";
            template.Security.AssociatedMemberGroup = "Site Members";
            template.Security.AssociatedVisitorGroup = "Dummy Visitors";
            // note: there are explicitly no SiteGroups defined

            using (var clientContext = TestCommon.CreateClientContext())
            {
                newCommSiteUrl = await CreateCommunicationSite(clientContext, "Dummy", true);
            }

            try
            {
                using (var ctx = TestCommon.CreateClientContext(newCommSiteUrl))
                {
                    ctx.Load(ctx.Web,
                        w => w.AssociatedOwnerGroup.Id,
                        w => w.AssociatedMemberGroup.Id,
                        w => w.AssociatedVisitorGroup.Id);
                    ctx.ExecuteQuery();
                    var oldOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var oldMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var oldVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;

                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
#if SP2019
                    await WaitForAsyncGroupTitleChangeWithTimeout(ctx);
#endif

                    ctx.Load(ctx.Web,
                        w => w.AssociatedOwnerGroup.Id,
                        w => w.AssociatedMemberGroup.Id,
                        w => w.AssociatedVisitorGroup.Id);
                    ctx.ExecuteQuery();
                    var newOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var newMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var newVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;

                    Assert.AreEqual(oldOwnerGroupId, newOwnerGroupId, "Unexpected new associated owner group");
                    Assert.AreEqual(oldMemberGroupId, newMemberGroupId, "Unexpected new associated member group");
                    Assert.AreEqual(oldVisitorGroupId, newVisitorGroupId, "Unexpected new associated visitor group");
                }
            }
            finally
            {
                using (var clientContext = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(clientContext);
#if !ONPREMISES
                    tenant.DeleteSiteCollection(newCommSiteUrl, false);
#else
                    tenant.DeleteSiteCollection(newCommSiteUrl);
#endif
                }
            }
        }

        // explicitly create new associated group with new title
        // for details see here: https://github.com/SharePoint/PnP-Sites-Core/pull/2174#issuecomment-487538551
        [TestMethod()]
        public async Task CanCreateNewAssociatedGroupsInNewScriptSite_CreateNewAssociatedGroup()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Site creation requires owner, so this will not yet work in app-only");

            this.CheckAndExcludeTestNotSupportedForAppOnlyTestting();

            var newCommSiteUrl = string.Empty;

            ProvisioningTemplate template = new ProvisioningTemplate();
            SiteGroup ownersGroup = new SiteGroup()
            {
                Title = string.Format("Test_New Group_{0}", DateTime.Now.Ticks),
            };
            template.Security.SiteGroups.Add(ownersGroup);
            template.Security.AssociatedOwnerGroup = ownersGroup.Title;

            using (var clientContext = TestCommon.CreateClientContext())
            {
                newCommSiteUrl = await CreateCommunicationSite(clientContext, "Dummy", true);
            }

            try
            {
                using (var ctx = TestCommon.CreateClientContext(newCommSiteUrl))
                {

                    LoadAssociatedOwnerGroupsData(ctx);
                    var oldOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var oldMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var oldVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;
                    var oldSiteGroupCount = ctx.Web.SiteGroups.Count;

                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
#if SP2019
                    await WaitForAsyncGroupTitleChangeWithTimeout(ctx);
#endif

                    LoadAssociatedOwnerGroupsData(ctx, true);
                    var newOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var newMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var newVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;
                    var newSiteGroupsCount = ctx.Web.SiteGroups.Count;
                    var group = ctx.Web.SiteGroups.First(sg => sg.Title == ownersGroup.Title);

                    Assert.AreEqual(group.Id, newOwnerGroupId, "Expected owners group to change");
                    Assert.AreEqual(oldMemberGroupId, newMemberGroupId, "Expected members group to stay the same");
                    Assert.AreEqual(oldVisitorGroupId, newVisitorGroupId, "Expected visitors group to stay the same");
                    Assert.AreEqual(oldSiteGroupCount + 1, newSiteGroupsCount, "Expected only one group to be created");

                    // and apply a second time to be sure this works as well
                    new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                    LoadAssociatedOwnerGroupsData(ctx, true);
                    Assert.AreEqual(newOwnerGroupId, ctx.Web.AssociatedOwnerGroup.Id, "Expected owners group to stay the same");
                    Assert.AreEqual(newMemberGroupId, ctx.Web.AssociatedMemberGroup.Id, "Expected members group to stay the same");
                    Assert.AreEqual(newVisitorGroupId, ctx.Web.AssociatedVisitorGroup.Id, "Expected visitors group to stay the same");
                    Assert.AreEqual(newSiteGroupsCount, ctx.Web.SiteGroups.Count, "Expected no new groups to be created");
                }
            }
            finally
            {
                using (var clientContext = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(clientContext);
#if !ONPREMISES
                    tenant.DeleteSiteCollection(newCommSiteUrl, false);
#else
                    tenant.DeleteSiteCollection(newCommSiteUrl);
#endif
                }
            }
        }

        // check mapping of existing group - before async group rename
        // for details see here: https://github.com/SharePoint/PnP-Sites-Core/pull/2174#issuecomment-487538551
        [TestMethod()]
        public async Task CanCreateNewAssociatedGroupsInNewScriptSite_BeforeAsyncRename()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Site creation requires owner, so this will not yet work in app-only");

            this.CheckAndExcludeTestNotSupportedForAppOnlyTestting();

            var newCommSiteUrl = string.Empty;
            var siteTitle = "Dummy";

            ProvisioningTemplate template = new ProvisioningTemplate();
            SiteGroup ownersGroup = new SiteGroup()
            {
                // this should map to existing group
                Title = "Site Owners"
            };
            template.Security.SiteGroups.Add(ownersGroup);
            template.Security.AssociatedOwnerGroup = ownersGroup.Title;

            using (var clientContext = TestCommon.CreateClientContext())
            {
                newCommSiteUrl = await CreateCommunicationSite(clientContext, siteTitle, true);
            }

            try
            {
                using (var ctx = TestCommon.CreateClientContext(newCommSiteUrl))
                {
                    LoadAssociatedOwnerGroupsData(ctx);
                    var oldOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var oldMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var oldVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;
                    var oldGroupCount = ctx.Web.SiteGroups.Count;

                    var parser = new TokenParser(ctx.Web, template);
                    // first provision - site titles are not yet in place
                    new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                    // wait for async rename
#if SP2019
                    await WaitForAsyncGroupTitleChangeWithTimeout(ctx);
#endif

                    LoadAssociatedOwnerGroupsData(ctx);
                    var newOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var newMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var newVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;
                    var newGroupCount = ctx.Web.SiteGroups.Count;

#if SP2019
                    Assert.AreEqual(oldOwnerGroupId, newOwnerGroupId, "Expected owners group to stay the same");
                    Assert.AreEqual(oldGroupCount, newGroupCount, "Expected no new groups to be created");
#else
                    Assert.AreNotEqual(oldOwnerGroupId, newOwnerGroupId, "Expected owners group is different since the added group 'Site Owners' was seen as a new group due to the waiting for async site creation completion in SPO");
                    Assert.AreEqual(oldGroupCount, newGroupCount - 1, "Expected owners group is different since the added group 'Site Owners' was seen as a new group due to the waiting for async site creation completion in SPO");
#endif
                    Assert.AreEqual(oldMemberGroupId, newMemberGroupId, "Expected members group to stay the same");
                    Assert.AreEqual(oldVisitorGroupId, newVisitorGroupId, "Expected visitors group to stay the same");
                }
            }
            finally
            {
                using (var clientContext = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(clientContext);
#if !ONPREMISES
                    tenant.DeleteSiteCollection(newCommSiteUrl, false);
#else
                    tenant.DeleteSiteCollection(newCommSiteUrl);
#endif
                }
            }
        }

        // check mapping to existing group - after async rename of site group titles
        // for details see here: https://github.com/SharePoint/PnP-Sites-Core/pull/2174#issuecomment-487538551
        [TestMethod()]
        public async Task CanCreateNewAssociatedGroupsInNewScriptSite_AfterAsyncRename()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Site creation requires owner, so this will not yet work in app-only");

            this.CheckAndExcludeTestNotSupportedForAppOnlyTestting();

            var newCommSiteUrl = string.Empty;
            var siteTitle = "Dummy";

            ProvisioningTemplate template = new ProvisioningTemplate();
            SiteGroup ownersGroup = new SiteGroup()
            {
                // this should map to existing group
                Title = siteTitle + " Owners",
            };
            template.Security.SiteGroups.Add(ownersGroup);
            template.Security.AssociatedOwnerGroup = ownersGroup.Title;

            using (var clientContext = TestCommon.CreateClientContext())
            {
                newCommSiteUrl = await CreateCommunicationSite(clientContext, siteTitle, true);
            }

            try
            {
                using (var ctx = TestCommon.CreateClientContext(newCommSiteUrl))
                {
                    LoadAssociatedOwnerGroupsData(ctx);
                    var oldOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var oldMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var oldVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;
                    var oldGroupCount = ctx.Web.SiteGroups.Count;

                    var parser = new TokenParser(ctx.Web, template);
                    // wait for async rename
#if SP2019
                    await WaitForAsyncGroupTitleChangeWithTimeout(ctx);
#endif
                    // now provision - new site titles are in place
                    new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                    LoadAssociatedOwnerGroupsData(ctx);
                    var newOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var newMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var newVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;
                    var newGroupCount = ctx.Web.SiteGroups.Count;

                    Assert.AreEqual(oldOwnerGroupId, newOwnerGroupId, "Expected owners group to stay the same");
                    Assert.AreEqual(oldMemberGroupId, newMemberGroupId, "Expected members group to stay the same");
                    Assert.AreEqual(oldVisitorGroupId, newVisitorGroupId, "Expected visitors group to stay the same");
                    Assert.AreEqual(oldGroupCount, newGroupCount, "Expected no new groups to be created");
                }
            }
            finally
            {
                using (var clientContext = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(clientContext);
#if !ONPREMISES
                    tenant.DeleteSiteCollection(newCommSiteUrl, false);
#else
                    tenant.DeleteSiteCollection(newCommSiteUrl);
#endif
                }
            }
        }

        // map owner groups to existing groups in the site
        // for details see here: https://github.com/SharePoint/PnP-Sites-Core/pull/2174#issuecomment-487538551
        [TestMethod()]
        public async Task CanMapToExistingGroupsInNewScriptSite()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Site creation requires owner, so this will not yet work in app-only");

            this.CheckAndExcludeTestNotSupportedForAppOnlyTestting();

            var newCommSiteUrl = string.Empty;

            ProvisioningTemplate template = new ProvisioningTemplate();
            // check setting of owner group with a configuration that would also create the group
            var ownerGroupTitle = string.Format("Test_New OWNER Group_{0}", DateTime.Now.Ticks);
            SiteGroup ownerGroup = new SiteGroup()
            {
                Title = ownerGroupTitle
            };
            template.Security.AssociatedOwnerGroup = ownerGroup.Title;
            template.Security.SiteGroups.Add(ownerGroup);

            // check setting of member group with a configuration that only works if the group already exists in the web
            var memberGroupTitle = string.Format("Test_New MEMBER Group_{0}", DateTime.Now.Ticks);
            template.Security.AssociatedMemberGroup = memberGroupTitle;

            using (var clientContext = TestCommon.CreateClientContext())
            {
                newCommSiteUrl = await CreateCommunicationSite(clientContext, "Dummy", true);
            }

            try
            {
                using (var ctx = TestCommon.CreateClientContext(newCommSiteUrl))
                {
                    // pre-create owner and member group to test mapping
                    ctx.Web.SiteGroups.Add(new GroupCreationInformation() { Title = ownerGroupTitle });
                    ctx.Web.SiteGroups.Add(new GroupCreationInformation() { Title = memberGroupTitle });
                    ctx.ExecuteQueryRetry();
                    ctx.Load(ctx.Web, w => w.SiteGroups.Include(
                        g => g.Id,
                        g => g.Title),
                        w => w.AssociatedVisitorGroup);
                    ctx.ExecuteQueryRetry();
                    var newOwnerGroup = ctx.Web.SiteGroups.First(g => g.Title == ownerGroupTitle);
                    var newMemberGroup = ctx.Web.SiteGroups.First(g => g.Title == memberGroupTitle);
                    var oldAssociatedVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;

                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectSiteSecurity().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                    LoadAssociatedOwnerGroupsData(ctx, true);
                    var associatedOwnerGroupId = ctx.Web.AssociatedOwnerGroup.Id;
                    var associatedMemberGroupId = ctx.Web.AssociatedMemberGroup.Id;
                    var associatedVisitorGroupId = ctx.Web.AssociatedVisitorGroup.Id;

                    Assert.AreEqual(newOwnerGroup.Id, associatedOwnerGroupId, "Expected owners group to change");
                    Assert.AreEqual(newMemberGroup.Id, associatedMemberGroupId, "Expected members group to change");
                    Assert.AreEqual(oldAssociatedVisitorGroupId, associatedVisitorGroupId, "Expected visitors group to stay the same");
                }
            }
            finally
            {
                using (var clientContext = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(clientContext);
#if !ONPREMISES
                    tenant.DeleteSiteCollection(newCommSiteUrl, false);
#else
                    tenant.DeleteSiteCollection(newCommSiteUrl);
#endif
                }
            }
        }

        private async Task<string> CreateCommunicationSite(ClientContext clientContext, string newSiteTitle, bool allowScripts = false)
        {
            var communicationSiteGuid = Guid.NewGuid().ToString("N");
            var baseUri = new Uri(clientContext.Url);
            var baseUrl = $"{baseUri.Scheme}://{baseUri.Host}:{baseUri.Port}";
            var newCommSiteUrl = $"{baseUrl}/sites/site{communicationSiteGuid}";

            // create new site to apply template to
            var siteCollectionCreationInformation = new Core.Sites.CommunicationSiteCollectionCreationInformation()
            {
                Url = $"{baseUrl}/sites/site{communicationSiteGuid}",
                SiteDesign = Core.Sites.CommunicationSiteDesign.Blank,
                Title = newSiteTitle,
                Lcid = 1033
            };

            // TODO: Owner set the owner
            /*
            if (clientContext.IsAppOnly()
                && string.IsNullOrEmpty(siteCollectionCreationInformation.Owner)
                && !string.IsNullOrEmpty(TestCommon.DefaultSiteOwner))
            {
                siteCollectionCreationInformation.Owner = TestCommon.DefaultSiteOwner;
            }
            */
            var commResults = await clientContext.CreateSiteAsync(siteCollectionCreationInformation);
            Assert.IsNotNull(commResults);

            if (allowScripts)
            {
                AllowScripting(newCommSiteUrl, true);
            }

            return newCommSiteUrl;
        }

        private void AllowScripting(string absoluteUrl, bool allow)
        {
            using (var adminCtx = TestCommon.CreateTenantClientContext())
            {
                var tenant = new Tenant(adminCtx);
                tenant.SetSiteProperties(absoluteUrl, noScriptSite: !allow);
            }
        }

        private async Task WaitForAsyncGroupTitleChangeWithTimeout(ClientContext ctx)
        {
            const int maxWaitMs = 4 /*minutes*/ * 60 * 1000;
            const int checkIntervalMs = 5000;
            const string tempTitleToCheckFor = "Site Owners";
            var waitMs = 0;
            // wait for async group title rename action to complete (can take ~2 minutes, is done by SharePoint)
            bool timeout = false;
            do
            {
                await Task.Delay(checkIntervalMs);
                waitMs += checkIntervalMs;
                ctx.Load(ctx.Web,
                    w => w.AssociatedOwnerGroup.Title);
                ctx.ExecuteQueryRetry();
                if (waitMs > maxWaitMs)
                {
                    timeout = true;
                    break;
                }
            } while (ctx.Web.AssociatedOwnerGroup.Title == tempTitleToCheckFor);
            if (timeout)
            {
                Assert.Fail("Waiting for async group title change timed out. Try increasing the maxWaitMs.");
            }
        }
#endif
        private void CheckAndExcludeTestNotSupportedForAppOnlyTestting()
        {
#if SP2019
            // TODO: Check SP2019 Support for Owner property in futures cumulative updates (After CU 2020-03)
            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive(
                    "Test that require creation of communication site collection are not supported in app-only for SP2019."
                    + Environment.NewLine
                    + " " + "An System.Exception: { 'SiteStatus':3,'SiteUrl':''} will be thrown."
                    + " " + "The Owner cannot be set in an app-only context. Cause SharePoint 2019 server side does not support the owner property yet."                    
                    + " " + "The property 'Owner' does not exist on type 'Microsoft.SharePoint.Portal.SPSiteCreationRequest'. Make sure to only use property names that are defined by the type."
                    + Environment.NewLine
                    + " " + "In HighTrust szenarios you can set the 'HighTrustBehalfOfUserLoginName'. See loginName parameter of AuthenticationManager.GetHighTrustCertificateAppAuthenticatedContext(string siteUrl, string clientId, string certificatePath, string certificatePassword, string certificateIssuerId, string loginName)."
                    ); ;

            }
#endif
        }

        private void LoadAssociatedOwnerGroupsData(ClientContext ctx, bool loadAllGroupsTitles = false)
        {
            if (!loadAllGroupsTitles)
            {
                ctx.Load(ctx.Web,
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id,
                    w => w.SiteGroups.Include(
                        sg => sg.Id
                        ));
            } else
            {
                ctx.Load(ctx.Web,
                    w => w.AssociatedOwnerGroup.Id,
                    w => w.AssociatedMemberGroup.Id,
                    w => w.AssociatedVisitorGroup.Id,
                    w => w.SiteGroups.Include(
                        sg => sg.Id,
                        sg => sg.Title
                        ));
            }
            ctx.ExecuteQueryRetry();
        }
    }
}
