using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Tests;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client.Tests
{
    [TestClass]
    public class SecurityExtensionsTests
    {
        private readonly string _testGroupName = "Group_" + Guid.NewGuid();
        private string _userLogin;

        #region Test initialize and cleanup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            // delete all the sub sites of 4 characters long...cleanup for potentially failed executions of
            // AddPermissionLevelToGroupSubSiteTest and RemovePermissionLevelFromGroupSubSiteTest tests
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var subWebs = clientContext.Web.Webs;
                clientContext.Load(subWebs, wc => wc.Include(w => w.ServerRelativeUrl));
                clientContext.ExecuteQueryRetry();

                for (int i = subWebs.Count - 1; i >= 0; i--)
                {
                    if (subWebs[i].ServerRelativeUrl.Split('/')[3].Length == 4)
                    {
                        try
                        {
                            subWebs[i].DeleteObject();
                            clientContext.ExecuteQueryRetry();
                        }
                        catch { }
                    }
                }
            }
        }

        [TestInitialize]
        public void Initialize()
        {

#if !ONPREMISES
            _userLogin = TestCommon.AppSetting("SPOUserName");
            if (TestCommon.AppOnlyTesting())
            {
                using (var clientContext = TestCommon.CreateClientContext())
                {
                    List<UserEntity> admins = clientContext.Web.GetAdministrators();
                    _userLogin = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[2];
                }
            }
#else
            _userLogin = String.Format(@"{0}\{1}", TestCommon.AppSetting("OnPremDomain"), TestCommon.AppSetting("OnPremUserName"));            
            if (TestCommon.AppOnlyTesting())
            {
                using (var clientContext = TestCommon.CreateClientContext())
                {
                    List<UserEntity> admins = clientContext.Web.GetAdministrators();
                    _userLogin = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[1];
                }
            }
#endif

            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                clientContext.Web.AddGroup(_testGroupName, "", true, true);
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                clientContext.Web.RemoveGroup(_testGroupName);
                if (clientContext.Web.GroupExists("Test Group"))
                {
                    clientContext.Web.RemoveGroup("Test Group");
                }
            }
        }
        #endregion

        #region Administrator tests
        [TestMethod]
        public void GetAdministratorsTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                Assert.IsTrue(clientContext.Web.GetAdministrators().Any(), "No administrators returned");
            }
        }

        [TestMethod]
        public void AddAdministratorsTest()
        {
            // Difficult to test on a developer (MSDN) tenant, as there is only one user allowed.
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                // Count admins
                int initialCount = clientContext.Web.GetAdministrators().Count;
                #if !ONPREMISES
                var userEntity = new UserEntity {LoginName = _userLogin, Email = _userLogin};
                #else
                var userEntity = new UserEntity { LoginName = _userLogin };
                #endif
                clientContext.Web.AddAdministrators(new List<UserEntity> {userEntity}, false);

                List<UserEntity> admins = clientContext.Web.GetAdministrators();
                bool found = false;
                foreach(var admin in admins) 
                {                    
                    string adminLoginName = admin.LoginName;
                    String[] parts = adminLoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length > 1)
                    {
                        adminLoginName = parts[2];
                    }
                    
                    if (adminLoginName.Equals(_userLogin, StringComparison.InvariantCultureIgnoreCase))
                    {
                        found = true;
                        break;
                    }
                }
                Assert.IsTrue(found);

                // Assumes that we're on a dev tenant, and that the existing sitecol admin is the same as the user being added.
                clientContext.Web.RemoveAdministrator(userEntity);
            }
        }
        #endregion

        #region Group tests
        [TestMethod]
        public void AddGroupTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                // Test
                Group group = clientContext.Web.AddGroup("Test Group", "Test Description", true);
                Assert.IsInstanceOfType(group, typeof (Group), "Group object returned not of correct type");
                Assert.IsTrue(group.Title == "Test Group", "Group not created with correct title");

                // Cleanup
                if (group != null)
                {
                    clientContext.Web.RemoveGroup(group);
                }
            }
        }

        [TestMethod]
        public void GroupExistsTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                bool groupExists = clientContext.Web.GroupExists(_testGroupName);
                Assert.IsTrue(groupExists);

                groupExists = clientContext.Web.GroupExists(_testGroupName + "987654321654367");
                Assert.IsFalse(groupExists);
            }
        }

        #endregion

        #region Permission level tests
        [TestMethod]
        public void AddPermissionLevelToGroupTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                // Test
                clientContext.Web.AddPermissionLevelToGroup(_testGroupName, RoleType.Contributor, false);

                //Get Group
                Group group = clientContext.Web.SiteGroups.GetByName(_testGroupName);
                clientContext.ExecuteQueryRetry();

                //Assert
                Assert.IsTrue(CheckPermissionOnPrinciple(clientContext.Web, group, RoleType.Contributor));
            }
        }

		[TestMethod]
		public void AddPermissionLevelToGroupSubSiteTest()
		{
			using (ClientContext clientContext = TestCommon.CreateClientContext())
			{
				//Arrange
				var subSite = CreateTestTeamSubSite(clientContext.Web);

                subSite.EnsureProperties(s => s.HasUniqueRoleAssignments);
				
				if (!subSite.HasUniqueRoleAssignments)
				{
					subSite.BreakRoleInheritance(false, true);
				}

				// Test
				subSite.AddPermissionLevelToGroup(_testGroupName, RoleType.Contributor, false);

				//Get Group
				Group group = subSite.SiteGroups.GetByName(_testGroupName);
				clientContext.ExecuteQueryRetry();

				//Assert
				Assert.IsTrue(CheckPermissionOnPrinciple(subSite, group, RoleType.Contributor));

				//Teardown
				subSite.DeleteObject();
				clientContext.ExecuteQueryRetry();
			}
		}

		[TestMethod]
		public void AddPermissionLevelToGroupListTest()
		{
			using (ClientContext clientContext = TestCommon.CreateClientContext())
			{
				//Arrange
				var list = clientContext.Web.CreateList(ListTemplateType.GenericList, GetRandomString(), false);

                list.EnsureProperties(l => l.HasUniqueRoleAssignments);
                
				if (!list.HasUniqueRoleAssignments)
				{
					list.BreakRoleInheritance(false, true);
				}

				// Test
				list.AddPermissionLevelToGroup(_testGroupName, RoleType.Contributor, false);

				//Get Group
				Group group = list.ParentWeb.SiteGroups.GetByName(_testGroupName);
				clientContext.ExecuteQueryRetry();

				//Assert
				Assert.IsTrue(CheckPermissionOnPrinciple(list, group, RoleType.Contributor));

				//Teardown
				list.DeleteObject();
				clientContext.ExecuteQueryRetry();
			}
		}

		[TestMethod]
		public void AddPermissionLevelToGroupListItemTest()
		{
			using (ClientContext clientContext = TestCommon.CreateClientContext())
			{
				//Arrange
				var list = clientContext.Web.CreateList(ListTemplateType.GenericList, GetRandomString(), false);
				var item = list.AddItem(new ListItemCreationInformation());
				item["Title"] = "Test";
				item.Update();
				clientContext.Load(item);
				clientContext.ExecuteQueryRetry();

                item.EnsureProperties(i => i.HasUniqueRoleAssignments);
				
				if (!item.HasUniqueRoleAssignments)
				{
					item.BreakRoleInheritance(false, true);
				}

				//Get Group
				Group group = list.ParentWeb.SiteGroups.GetByName(_testGroupName);
				clientContext.ExecuteQueryRetry();

				// Test
				item.AddPermissionLevelToPrincipal(group, RoleType.Contributor, false);

				//Assert
				Assert.IsTrue(CheckPermissionOnPrinciple(item, group, RoleType.Contributor));

				//Teardown
				list.DeleteObject();
				clientContext.ExecuteQueryRetry();
			}
		}

		[TestMethod]
		public void RemovePermissionLevelFromGroupSubSiteTest()
		{
			using (ClientContext clientContext = TestCommon.CreateClientContext())
			{
				//Arrange
				var subSite = CreateTestTeamSubSite(clientContext.Web);


                subSite.EnsureProperties(s => s.HasUniqueRoleAssignments);
				
				if (!subSite.HasUniqueRoleAssignments)
				{
					subSite.BreakRoleInheritance(true, true);
				}

				//Get Group
				Group group = subSite.SiteGroups.GetByName(_testGroupName);
				clientContext.ExecuteQueryRetry();

				subSite.AddPermissionLevelToPrincipal(group, RoleType.Contributor);
				subSite.AddPermissionLevelToPrincipal(group, RoleType.Editor);

				// Test
				subSite.RemovePermissionLevelFromPrincipal(group, RoleType.Contributor);

				//Assert
				Assert.IsFalse(CheckPermissionOnPrinciple(subSite, group, RoleType.Contributor));

				//Teardown
				subSite.DeleteObject();
				clientContext.ExecuteQueryRetry();
			}
		}

        [TestMethod]
        public void AddPermissionLevelByRoleDefToGroupTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                // Test
                clientContext.Web.AddPermissionLevelToGroup(_testGroupName, "Approve", false);

                //Get Group
                Group group = clientContext.Web.SiteGroups.GetByName(_testGroupName);
                clientContext.ExecuteQueryRetry();

                //Assert 
                Assert.IsTrue(CheckPermissionOnPrinciple(clientContext.Web, group, "Approve"));
            }
        }

        [TestMethod]
        public void AddPermissionLevelToUserTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                Web web = clientContext.Web;

                var roleType = RoleType.Contributor;

                //Setup: Make sure permission does not already exist
                web.RemovePermissionLevelFromUser(_userLogin, roleType);

                //Add Permission
                web.AddPermissionLevelToUser(_userLogin, roleType);

                //Get User
                User user = web.EnsureUser(_userLogin);
                clientContext.ExecuteQueryRetry();

                //Assert
                Assert.IsTrue(CheckPermissionOnPrinciple(web, user, roleType));

                //Teardown: Expicitly remove given permission. 
                web.RemovePermissionLevelFromUser(_userLogin, roleType);
            }
        }

        [TestMethod]
        public void AddPermissionLevelToUserTestByRoleDefTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                Web web = clientContext.Web;
				
                //Setup: Make sure permission does not already exist
                web.RemovePermissionLevelFromUser(_userLogin, "Approve");

                //Add Permission
                web.AddPermissionLevelToUser(_userLogin, "Approve");

                //Get User
                User user = web.EnsureUser(_userLogin);
                clientContext.ExecuteQueryRetry();

                //Assert
                Assert.IsTrue(CheckPermissionOnPrinciple(web, user, "Approve"));

                //Teardown: Expicitly remove given permission. 
                web.RemovePermissionLevelFromUser(_userLogin, "Approve");
            }
        }

		[TestMethod]
		public void AddSamePermissionLevelTwiceToGroupTest()
		{
			using (ClientContext clientContext = TestCommon.CreateClientContext())
			{
				// Test
				clientContext.Web.AddPermissionLevelToGroup(_testGroupName, RoleType.Contributor, true);
				clientContext.Web.AddPermissionLevelToGroup(_testGroupName, RoleType.Contributor, false);

				//Get Group
				Group group = clientContext.Web.SiteGroups.GetByName(_testGroupName);
				clientContext.ExecuteQueryRetry();

				//Assert
				Assert.IsTrue(CheckPermissionOnPrinciple(clientContext.Web, group, RoleType.Contributor));
			}
		}

        #endregion

        #region Reader access tests
#if !ONPREMISES
        [TestMethod]
        public void AddReaderAccessToEveryoneExceptExternalsTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                // Setup
                User userIdentity = null;

                // Test
                userIdentity = clientContext.Web.AddReaderAccess();

                Assert.IsNotNull(userIdentity, "No user added");
                User existingUser = clientContext.Web.AssociatedVisitorGroup.Users.GetByLoginName(userIdentity.LoginName);

                Assert.IsNotNull(existingUser, "No user returned");
                Assert.IsInstanceOfType(existingUser, typeof (User), "Object returned not of correct type");

                // Cleanup
                if (existingUser != null)
                {
                    clientContext.Web.AssociatedVisitorGroup.Users.Remove(existingUser);
                    clientContext.Web.AssociatedVisitorGroup.Update();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }
#endif

        [TestMethod]
        public void AddReaderAccessToEveryoneTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                // Setup
                string userIdentity = "c:0(.s|true";

                // Test
                clientContext.Web.AddReaderAccess(BuiltInIdentity.Everyone);

                User existingUser = clientContext.Web.AssociatedVisitorGroup.Users.GetByLoginName(userIdentity);
                Assert.IsNotNull(existingUser, "No user returned");
                Assert.IsInstanceOfType(existingUser, typeof (User), "Object returned not of correct type");

                // Cleanup
                if (existingUser != null)
                {
                    clientContext.Web.AssociatedVisitorGroup.Users.Remove(existingUser);
                    clientContext.Web.AssociatedVisitorGroup.Update();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }
        #endregion

        #region Get all unique role assignments tests

        [TestMethod]
        public void GetAllUniqueRoleAssignmentsTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                var assignments = clientContext.Web.GetAllUniqueRoleAssignments();
                Assert.AreNotEqual(null, assignments);
                Assert.AreNotEqual(0, assignments.Count());
                foreach (var item in assignments)
                {
                    Trace.WriteLine(item);
                }                
            }
        }

        #endregion

        #region helper methods
        private bool CheckPermissionOnPrinciple(SecurableObject securableObject, Principal principle, RoleType roleType)
        {
            //Get Roles for the User
            RoleDefinitionBindingCollection roleDefinitionBindingCollection =
                securableObject.RoleAssignments.GetByPrincipal(principle).RoleDefinitionBindings;
            securableObject.Context.Load(roleDefinitionBindingCollection);
            securableObject.Context.ExecuteQueryRetry();

            //Check if assigned role is found
            bool roleExists = false;
            foreach (RoleDefinition rd in roleDefinitionBindingCollection)
            {
                if (rd.RoleTypeKind == roleType)
                {
                    roleExists = true;
                }
            }

            return roleExists;
        }

        private bool CheckPermissionOnPrinciple(SecurableObject securableObject, Principal principle, string roleDefinitionName)
        {
            //Get Roles for the User
            RoleDefinitionBindingCollection roleDefinitionBindingCollection =
                securableObject.RoleAssignments.GetByPrincipal(principle).RoleDefinitionBindings;
            securableObject.Context.Load(roleDefinitionBindingCollection);
            securableObject.Context.ExecuteQueryRetry();

            //Check if assigned role is found
            bool roleExists = false;
            foreach (RoleDefinition rd in roleDefinitionBindingCollection)
            {
                if (rd.Name == roleDefinitionName)
                {
                    roleExists = true;
                }
            }

            return roleExists;
        
		}

	    private Web CreateTestTeamSubSite(Web parentWeb)
	    {
		    var siteUrl = GetRandomString();
		    var webInfo = new WebCreationInformation
		    {
			    Title = siteUrl,
			    Url = siteUrl,
			    Description = siteUrl,
			    Language = 1033,
			    UseSamePermissionsAsParentSite = true,
			    WebTemplate = "STS#0"
		    };

		    var web = parentWeb.Webs.Add(webInfo);
			parentWeb.Context.Load(web);
			parentWeb.Context.ExecuteQueryRetry();

            using (var ctxTestTeamSubSite = parentWeb.Context.Clone(TestCommon.DevSiteUrl + "/" + siteUrl))
            {
                return ctxTestTeamSubSite.Web;
            }
	    }

	    private string GetRandomString()
	    {
			var chars = "abcdefghijklmnopqrstuvwxyz";
			var random = new Random();
			return new string(Enumerable.Repeat(chars, 4).Select(s => s[random.Next(s.Length)]).ToArray());
	    }
        #endregion
    }
}