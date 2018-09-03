using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Utilities;
using User = OfficeDevPnP.Core.Framework.Provisioning.Model.User;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectSiteSecurityTests
    {

        private List<UserEntity> admins;

        [TestInitialize]
        public void Initialize()
        {

            using (var ctx = TestCommon.CreateClientContext())
            {
                admins = ctx.Web.GetAdministrators();
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var memberGroup = ctx.Web.AssociatedMemberGroup;
                ctx.Load(memberGroup);
                ctx.ExecuteQueryRetry();
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
    }
}
