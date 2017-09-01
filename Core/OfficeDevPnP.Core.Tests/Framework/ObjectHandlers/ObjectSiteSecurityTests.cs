using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
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
    }
}
