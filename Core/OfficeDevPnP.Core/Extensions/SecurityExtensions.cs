using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
#if NETSTANDARD2_0
using System.Net;
#endif
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.AppModelExtensions;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// This manager class holds security related methods
    /// </summary>
    public static partial class SecurityExtensions
    {
        #region Site collection administrator management
        /// <summary>
        /// Get a list of site collection administrators
        /// </summary>
        /// <param name="web">Site to operate on</param>
        /// <returns>List of <see cref="OfficeDevPnP.Core.Entities.UserEntity"/> objects</returns>
        public static List<UserEntity> GetAdministrators(this Web web)
        {
            var users = web.SiteUsers;
            web.Context.Load(users);
            web.Context.ExecuteQueryRetry();

            List<UserEntity> admins = new List<UserEntity>();

            foreach (var u in users)
            {
                if (u.IsSiteAdmin)
                {
                    admins.Add(new UserEntity()
                    {
                        Title = u.Title,
                        LoginName = u.LoginName,
                        Email = u.Email,
                    });
                }
            }

            return admins;
        }

        /// <summary>
        /// Add a site collection administrator to a site collection
        /// </summary>
        /// <param name="web">Site to operate on</param>
        /// <param name="adminLogins">Array of admins loginnames to add</param>
        /// <param name="addToOwnersGroup">Optionally the added admins can also be added to the Site owners group</param>
        public static void AddAdministrators(this Web web, List<UserEntity> adminLogins, bool addToOwnersGroup = false)
        {
            var users = web.SiteUsers;
            web.Context.Load(users);

            foreach (var admin in adminLogins)
            {
                UserCreationInformation newAdmin = new UserCreationInformation();

                newAdmin.LoginName = admin.LoginName;
                //User addedAdmin = users.Add(newAdmin);
                User addedAdmin = web.EnsureUser(newAdmin.LoginName);
                web.Context.Load(addedAdmin);
                web.Context.ExecuteQueryRetry();

                //now that the user exists in the context, update to be an admin
                addedAdmin.IsSiteAdmin = true;
                addedAdmin.Update();

                if (addToOwnersGroup)
                {
                    web.AssociatedOwnerGroup.Users.AddUser(addedAdmin);
                    web.AssociatedOwnerGroup.Update();
                }
                web.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Removes an administrators from the site collection
        /// </summary>
        /// <param name="web">Site to operate on</param>
        /// <param name="admin"><see cref="OfficeDevPnP.Core.Entities.UserEntity"/> that describes the admin to be removed</param>
        public static void RemoveAdministrator(this Web web, UserEntity admin)
        {
            var users = web.SiteUsers;
            web.Context.Load(users);
            web.Context.ExecuteQueryRetry();

            var adminToRemove = users.FirstOrDefault(u => String.Equals(u.LoginName, admin.LoginName, StringComparison.CurrentCultureIgnoreCase));
            if (adminToRemove != null && adminToRemove.IsSiteAdmin)
            {
                adminToRemove.IsSiteAdmin = false;
                adminToRemove.Update();
                web.Context.ExecuteQueryRetry();
            }

        }


        #endregion

        #region Permissions management
        /// <summary>
        /// Add read access to the group "Everyone except external users".
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        public static User AddReaderAccess(this Web web)
        {
            return AddReaderAccessImplementation(web, BuiltInIdentity.EveryoneButExternalUsers);
        }

        /// <summary>
        /// Add read access to the group "Everyone except external users".
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="user">Built in user to add to the visitors group</param>
        public static User AddReaderAccess(this Web web, BuiltInIdentity user)
        {
            return AddReaderAccessImplementation(web, user);
        }

        private static User AddReaderAccessImplementation(Web web, BuiltInIdentity user)
        {
            switch (user)
            {
                case BuiltInIdentity.Everyone:
                    {
                        const string userIdentity = "c:0(.s|true";
                        User spReader = web.EnsureUser(userIdentity);
                        web.Context.Load(spReader);
                        web.Context.ExecuteQueryRetry();

                        web.AssociatedVisitorGroup.Users.AddUser(spReader);
                        web.AssociatedVisitorGroup.Update();
                        web.Context.ExecuteQueryRetry();
                        return spReader;
                    }
                case BuiltInIdentity.EveryoneButExternalUsers:
                    {
                        User spReader = null;
                        try
                        {
                            // New tenant
                            string userIdentity =
                                $"c:0-.f|rolemanager|spo-grid-all-users/{web.GetAuthenticationRealm()}";
                            spReader = web.EnsureUser(userIdentity);
                            web.Context.Load(spReader);
                            web.Context.ExecuteQueryRetry();
                        }
                        catch (ServerException)
                        {
                            // old tenant?
                            string userIdentity = string.Empty;

                            userIdentity = web.GetEveryoneExceptExternalUsersClaimName();

                            if (!string.IsNullOrEmpty(userIdentity))
                            {
                                spReader = web.EnsureUser(userIdentity);
                                web.Context.Load(spReader);
                                web.Context.ExecuteQueryRetry();
                            }
                            else
                            {
                                throw new Exception("Language currently not supported");
                            }
                        }
                        web.AssociatedVisitorGroup.Users.AddUser(spReader);
                        web.AssociatedVisitorGroup.Update();
                        web.Context.ExecuteQueryRetry();
                        return spReader;
                    }
            }

            return null;
        }

        /// <summary>
        /// Returns the correct value of the "Everyone except external users" string value
        /// </summary>
        /// <param name="web">Web to get the language from</param>
        /// <returns>String in correct translation</returns>
        public static string GetEveryoneExceptExternalUsersClaimName(this Web web)
        {
            string userIdentity = string.Empty;

            web.EnsureProperty(p => p.Language);

            switch (web.Language)
            {
                case 1025: // Arabic
                    userIdentity = "الجميع باستثناء المستخدمين الخارجيين";
                    break;
                case 1069: // Basque
                    userIdentity = "Guztiak kanpoko erabiltzaileak izan ezik";
                    break;
                case 1026: // Bulgarian
                    userIdentity = "Всички освен външни потребители";
                    break;
                case 1027: // Catalan
                    userIdentity = "Tothom excepte els usuaris externs";
                    break;
                case 2052: // Chinese (Simplified)
                    userIdentity = "除外部用户外的任何人";
                    break;
                case 1028: // Chinese (Traditional)
                    userIdentity = "外部使用者以外的所有人";
                    break;
                case 1050: // Croatian
                    userIdentity = "Svi osim vanjskih korisnika";
                    break;
                case 1029: // Czech
                    userIdentity = "Všichni kromě externích uživatelů";
                    break;
                case 1030: // Danish
                    userIdentity = "Alle undtagen eksterne brugere";
                    break;
                case 1043: // Dutch
                    userIdentity = "Iedereen behalve externe gebruikers";
                    break;
                case 1033: // English
                    userIdentity = "Everyone except external users";
                    break;
                case 1061: // Estonian
                    userIdentity = "Kõik peale väliskasutajate";
                    break;
                case 1035: // Finnish
                    userIdentity = "Kaikki paitsi ulkoiset käyttäjät";
                    break;
                case 1036: // French
                    userIdentity = "Tout le monde sauf les utilisateurs externes";
                    break;
                case 1110: // Galician
                    userIdentity = "Todo o mundo excepto os usuarios externos";
                    break;
                case 1031: // German
                    userIdentity = "Jeder, außer externen Benutzern";
                    break;
                case 1032: // Greek
                    userIdentity = "Όλοι εκτός από εξωτερικούς χρήστες";
                    break;
                case 1037: // Hebrew
                    userIdentity = "כולם פרט למשתמשים חיצוניים";
                    break;
                case 1081: // Hindi
                    userIdentity = "बाह्य उपयोगकर्ताओं को छोड़कर सभी";
                    break;
                case 1038: // Hungarian
                    userIdentity = "Mindenki, kivéve külső felhasználók";
                    break;
                case 1057: // Indonesian
                    userIdentity = "Semua orang kecuali pengguna eksternal";
                    break;
                case 1040: // Italian
                    userIdentity = "Tutti tranne gli utenti esterni";
                    break;
                case 1041: // Japanese
                    userIdentity = "外部ユーザー以外のすべてのユーザー";
                    break;
                case 1087: // Kazakh
                    userIdentity = "Сыртқы пайдаланушылардан басқасының барлығы";
                    break;
                case 1042: // Korean
                    userIdentity = "외부 사용자를 제외한 모든 사람";
                    break;
                case 1062: // Latvian
                    userIdentity = "Visi, izņemot ārējos lietotājus";
                    break;
                case 1063: // Lithuanian
                    userIdentity = "Visi, išskyrus išorinius vartotojus";
                    break;
                case 1086: // Malay
                    userIdentity = "Semua orang kecuali pengguna luaran";
                    break;
                case 1044: // Norwegian (Bokmål)
                    userIdentity = "Alle bortsett fra eksterne brukere";
                    break;
                case 1045: // Polish
                    userIdentity = "Wszyscy oprócz użytkowników zewnętrznych";
                    break;
                case 1046: // Portuguese (Brazil)
                    userIdentity = "Todos exceto os usuários externos";
                    break;
                case 2070: // Portuguese (Portugal)
                    userIdentity = "Todos exceto os utilizadores externos";
                    break;
                case 1048: // Romanian
                    userIdentity = "Toată lumea, cu excepția utilizatorilor externi";
                    break;
                case 1049: // Russian
                    userIdentity = "Все, кроме внешних пользователей";
                    break;
                case 10266: // Serbian (Cyrillic, Serbia)
                    userIdentity = "Сви осим спољних корисника";
                    break;
                case 2074:// Serbian (Latin)
                    userIdentity = "Svi osim spoljnih korisnika";
                    break;
                case 1051:// Slovak
                    userIdentity = "Všetci okrem externých používateľov";
                    break;
                case 1060: // Slovenian
                    userIdentity = "Vsi razen zunanji uporabniki";
                    break;
                case 3082: // Spanish
                    userIdentity = "Todos excepto los usuarios externos";
                    break;
                case 1053: // Swedish
                    userIdentity = "Alla utom externa användare";
                    break;
                case 1054: // Thai
                    userIdentity = "ทุกคนยกเว้นผู้ใช้ภายนอก";
                    break;
                case 1055: // Turkish
                    userIdentity = "Dış kullanıcılar hariç herkes";
                    break;
                case 1058: // Ukranian
                    userIdentity = "Усі, крім зовнішніх користувачів";
                    break;
                case 1066: // Vietnamese
                    userIdentity = "Tất cả mọi người trừ người dùng bên ngoài";
                    break;
            }

            return userIdentity;
        }

        #endregion

#if !ONPREMISES
        #region External sharing management
        /// <summary>
        /// Get the external sharing settings for the provided site. Only works in Office 365 Multi-Tenant
        /// </summary>
        /// <param name="web">Tenant administration web</param>
        /// <param name="siteUrl">Site to get the sharing capabilities from</param>
        /// <returns>Sharing capabilities of the site collection</returns>
        public static string GetSharingCapabilitiesTenant(this Web web, Uri siteUrl)
        {
            if (siteUrl == null)
                throw new ArgumentNullException("siteUrl");

            Tenant tenant = new Tenant(web.Context);
            SiteProperties site = tenant.GetSitePropertiesByUrl(siteUrl.OriginalString, true);
            web.Context.Load(site);
            web.Context.ExecuteQueryRetry();
            return site.SharingCapability.ToString();
        }

        /// <summary>
        /// Returns a list all external users in your tenant
        /// </summary>
        /// <param name="web">Tenant administration web</param>
        /// <returns>A list of <see cref="OfficeDevPnP.Core.Entities.ExternalUserEntity"/> objects</returns>
        public static List<ExternalUserEntity> GetExternalUsersTenant(this Web web)
        {
            Tenant tenantAdmin = new Tenant(web.Context);
            Office365Tenant tenant = new Office365Tenant(web.Context);

            List<ExternalUserEntity> externalUsers = new List<ExternalUserEntity>();
            const int pageSize = 50;
            int position = 0;

            while (true)
            {
                var results = tenant.GetExternalUsers(position, pageSize, string.Empty, SortOrder.Ascending);
                web.Context.Load(results, r => r.UserCollectionPosition, r => r.TotalUserCount, r => r.ExternalUserCollection);
                web.Context.ExecuteQueryRetry();

                foreach (var externalUser in results.ExternalUserCollection)
                {
                    externalUsers.Add(new ExternalUserEntity()
                    {
                        DisplayName = externalUser.DisplayName,
                        AcceptedAs = externalUser.AcceptedAs,
                        InvitedAs = externalUser.InvitedAs,
                        UniqueId = externalUser.UniqueId,
                        InvitedBy = externalUser.InvitedBy,
                        WhenCreated = externalUser.WhenCreated,
                    });
                }

                position = results.UserCollectionPosition;

                if (position == -1 || position == results.TotalUserCount)
                {
                    break;
                }
            }

            return externalUsers;
        }


        /// <summary>
        /// Returns a list all external users for a given site that have at least the viewpages permission
        /// </summary>
        /// <param name="web">Tenant administration web</param>
        /// <param name="siteUrl">Url of the site fetch the external users for</param>
        /// <returns>A list of <see cref="OfficeDevPnP.Core.Entities.ExternalUserEntity"/> objects</returns>
        public static List<ExternalUserEntity> GetExternalUsersForSiteTenant(this Web web, Uri siteUrl)
        {
            if (siteUrl == null)
                throw new ArgumentNullException("siteUrl");

            Tenant tenantAdmin = new Tenant(web.Context);
            Office365Tenant tenant = new Office365Tenant(web.Context);
            Site site = tenantAdmin.GetSiteByUrl(siteUrl.OriginalString);
            web = site.RootWeb;

            List<ExternalUserEntity> externalUsers = new List<ExternalUserEntity>();
            const int pageSize = 50;
            int position = 0;

            while (true)
            {
                var results = tenant.GetExternalUsersForSite(siteUrl.OriginalString, position, pageSize, string.Empty, SortOrder.Ascending);
                web.Context.Load(results, r => r.UserCollectionPosition, r => r.TotalUserCount, r => r.ExternalUserCollection);
                web.Context.ExecuteQueryRetry();

                foreach (var externalUser in results.ExternalUserCollection)
                {

                    User user = web.SiteUsers.GetByEmail(externalUser.AcceptedAs);
                    web.Context.Load(user);
                    web.Context.ExecuteQueryRetry();

                    var permission = web.GetUserEffectivePermissions(user.LoginName);
                    web.Context.ExecuteQueryRetry();
                    var doesUserHavePermission = permission.Value.Has(PermissionKind.ViewPages);
                    if (doesUserHavePermission)
                    {
                        externalUsers.Add(new ExternalUserEntity()
                        {
                            DisplayName = externalUser.DisplayName,
                            AcceptedAs = externalUser.AcceptedAs,
                            InvitedAs = externalUser.InvitedAs,
                            UniqueId = externalUser.UniqueId,
                            InvitedBy = externalUser.InvitedBy,
                            WhenCreated = externalUser.WhenCreated,
                        });
                    }

                }

                position = results.UserCollectionPosition;

                if (position == -1 || position == results.TotalUserCount)
                {
                    break;
                }
            }

            return externalUsers;
        }

        #endregion
#endif

        #region Group management
        /// <summary>
        /// Returns the integer ID for a given group name
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="groupName">SharePoint group name</param>
        /// <returns>Integer group ID</returns>
        public static int GetGroupID(this Web web, string groupName)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            int groupID = 0;

            var manageMessageGroup = web.SiteGroups.GetByName(groupName);
            web.Context.Load(manageMessageGroup);
            web.Context.ExecuteQueryRetry();
            if (manageMessageGroup != null)
            {
                groupID = manageMessageGroup.Id;
            }

            return groupID;
        }

        /// <summary>
        /// Adds a group
        /// </summary>
        /// <param name="web">Site to add the group to</param>
        /// <param name="groupName">Name of the group</param>
        /// <param name="groupDescription">Description of the group</param>
        /// <param name="groupIsOwner">Sets the created group as group owner if true</param>
        /// <param name="updateAndExecuteQuery">Set to false to postpone the executequery call</param>
        /// <param name="onlyAllowMembersViewMembership">Set whether members are allowed to see group membership, defaults to false</param>
        /// <returns>The created group</returns>
        public static Group AddGroup(this Web web, string groupName, string groupDescription, bool groupIsOwner, bool updateAndExecuteQuery = true, bool onlyAllowMembersViewMembership = false)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            GroupCreationInformation groupCreationInformation = new GroupCreationInformation();
            groupCreationInformation.Title = groupName;
            groupCreationInformation.Description = groupDescription;
            Group group = web.SiteGroups.Add(groupCreationInformation);
            if (groupIsOwner)
            {
                group.Owner = group;
            }

            group.OnlyAllowMembersViewMembership = onlyAllowMembersViewMembership;
            group.Update();

            if (updateAndExecuteQuery)
            {
                web.Context.ExecuteQueryRetry();
            }

            return group;
        }

        /// <summary>
        /// Associate the provided groups as default owners, members or visitors groups. If a group is null then the
        /// association is not done
        /// </summary>
        /// <param name="web">Site to operate on</param>
        /// <param name="owners">Owners group</param>
        /// <param name="members">Members group</param>
        /// <param name="visitors">Visitors group</param>
        public static void AssociateDefaultGroups(this Web web, Group owners, Group members, Group visitors)
        {
            if (owners != null)
            {
                web.AssociatedOwnerGroup = owners;
                web.AssociatedOwnerGroup.Update();
            }
            if (members != null)
            {
                web.AssociatedMemberGroup = members;
                web.AssociatedMemberGroup.Update();
            }
            if (visitors != null)
            {
                web.AssociatedVisitorGroup = visitors;
                web.AssociatedVisitorGroup.Update();
            }

            web.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Adds a user to a group
        /// </summary>
        /// <param name="web">web to operate against</param>
        /// <param name="groupName">Name of the group</param>
        /// <param name="userLoginName">Loginname of the user</param>
        public static void AddUserToGroup(this Web web, string groupName, string userLoginName)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            //Ensure the user is known
            UserCreationInformation userToAdd = new UserCreationInformation();
            userToAdd.LoginName = userLoginName;
            User user = web.EnsureUser(userToAdd.LoginName);
            web.Context.Load(user);
            //web.Context.ExecuteQueryRetry();

            //Add the user to the group
            var group = web.SiteGroups.GetByName(groupName);
            web.Context.Load(group);
            web.Context.ExecuteQueryRetry();
            if (group != null)
            {
                web.AddUserToGroup(group, user);
            }
        }

        /// <summary>
        /// Adds a user to a group
        /// </summary>
        /// <param name="web">web to operate against</param>
        /// /// <param name="groupId">Id of the group</param>
        /// <param name="userLoginName">Login name of the user</param>
        public static void AddUserToGroup(this Web web, int groupId, string userLoginName)
        {
            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            Group group = web.SiteGroups.GetById(groupId);
            web.Context.Load(group);
            User user = web.EnsureUser(userLoginName);
            web.Context.ExecuteQueryRetry();

            if (user != null && group != null)
            {
                AddUserToGroup(web, group, user);
            }
        }

        /// <summary>
        /// Adds a user to a group
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="group">Group object representing the group</param>
        /// <param name="user">User object representing the user</param>
        public static void AddUserToGroup(this Web web, Group group, User user)
        {
            if (group == null)
                throw new ArgumentNullException("group");

            if (user == null)
                throw new ArgumentNullException("user");

            group.Users.AddUser(user);
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Adds a user to a group
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="group">Group object representing the group</param>
        /// <param name="userLoginName">Login name of the user</param>
        public static void AddUserToGroup(this Web web, Group group, string userLoginName)
        {
            if (group == null)
                throw new ArgumentNullException("group");

            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            User user = web.EnsureUser(userLoginName);
            web.Context.ExecuteQueryRetry();
            if (user != null)
            {
                group.Users.AddUser(user);
                web.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Add a permission level (e.g.Contribute, Reader,...) to a user
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="userLoginName">Loginname of the user</param>
        /// <param name="permissionLevel">Permission level to add</param>
        /// <param name="removeExistingPermissionLevels">Set to true to remove all other permission levels for that user</param>
        public static void AddPermissionLevelToUser(this SecurableObject securableObject, string userLoginName, RoleType permissionLevel, bool removeExistingPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            Web web = securableObject.GetAssociatedWeb();

            User user = web.EnsureUser(userLoginName);
            securableObject.AddPermissionLevelToPrincipal(user, permissionLevel, removeExistingPermissionLevels);
        }

        /// <summary>
        /// Add a role definition (e.g.Contribute, Read, Approve) to a user
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="userLoginName">Loginname of the user</param>
        /// <param name="roleDefinitionName">Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Hierarchy|Restricted Read. Use the correct name of the language of the root site you are using</param>
        /// <param name="removeExistingPermissionLevels">Set to true to remove all other permission levels for that user</param>
        public static void AddPermissionLevelToUser(this SecurableObject securableObject, string userLoginName, string roleDefinitionName, bool removeExistingPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            if (string.IsNullOrEmpty(roleDefinitionName))
                throw new ArgumentNullException("roleDefinitionName");

            Web web = securableObject.GetAssociatedWeb();

            User user = web.EnsureUser(userLoginName);
            securableObject.AddPermissionLevelToPrincipal(user, roleDefinitionName, removeExistingPermissionLevels);
        }

        /// <summary>
        /// Add a permission level (e.g.Contribute, Reader,...) to a group
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="groupName">Name of the group</param>
        /// <param name="permissionLevel">Permission level to add</param>
        /// <param name="removeExistingPermissionLevels">Set to true to remove all other permission levels for that group</param>
        public static void AddPermissionLevelToGroup(this SecurableObject securableObject, string groupName, RoleType permissionLevel, bool removeExistingPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            Web web = securableObject.GetAssociatedWeb();

            var group = web.SiteGroups.GetByName(groupName);

            securableObject.AddPermissionLevelToPrincipal(group, permissionLevel, removeExistingPermissionLevels);
        }

        /// <summary>
        /// Add a permission level (e.g.Contribute, Reader,...) to a group
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="principal">Principal to add permission to</param>
        /// <param name="permissionLevel">Permission level to add</param>
        /// <param name="removeExistingPermissionLevels">Set to true to remove all other permission levels for that group</param>
        public static void AddPermissionLevelToPrincipal(this SecurableObject securableObject, Principal principal, RoleType permissionLevel, bool removeExistingPermissionLevels = false)
        {
            if (principal == null)
                throw new ArgumentNullException("principal");

            Web web = securableObject.GetAssociatedWeb();

            securableObject.Context.Load(principal);
            securableObject.Context.ExecuteQueryRetry();
            RoleDefinition roleDefinition = web.RoleDefinitions.GetByType(permissionLevel);
            securableObject.AddPermissionLevelImplementation(principal, roleDefinition, removeExistingPermissionLevels);
        }

        /// <summary>
        /// Add a role definition (e.g.Contribute, Read, Approve) to a group
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="groupName">Name of the group</param>
        /// <param name="roleDefinitionName">Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Hierarchy|Restricted Read. Use the correct name of the language of the root site you are using</param>
        /// <param name="removeExistingPermissionLevels">Set to true to remove all other permission levels for that group</param>
        public static void AddPermissionLevelToGroup(this SecurableObject securableObject, string groupName, string roleDefinitionName, bool removeExistingPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            if (string.IsNullOrEmpty(roleDefinitionName))
                throw new ArgumentNullException("roleDefinitionName");

            Web web = securableObject.GetAssociatedWeb();

            var group = web.SiteGroups.GetByName(groupName);

            securableObject.AddPermissionLevelToPrincipal(group, roleDefinitionName, removeExistingPermissionLevels);
        }

        /// <summary>
        /// Add a role definition (e.g.Contribute, Read, Approve) to a group
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="principal">Principal to add permission to</param>
        /// <param name="roleDefinitionName">Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Hierarchy|Restricted Read. Use the correct name of the language of the root site you are using</param>
        /// <param name="removeExistingPermissionLevels">Set to true to remove all other permission levels for that group</param>
        public static void AddPermissionLevelToPrincipal(this SecurableObject securableObject, Principal principal, string roleDefinitionName, bool removeExistingPermissionLevels = false)
        {
            if (principal == null)
                throw new ArgumentNullException("principal");

            if (string.IsNullOrEmpty(roleDefinitionName))
                throw new ArgumentNullException("roleDefinitionName");

            Web web = securableObject.GetAssociatedWeb();

            securableObject.Context.Load(principal);
            securableObject.Context.ExecuteQueryRetry();
            RoleDefinition roleDefinition = web.RoleDefinitions.GetByName(roleDefinitionName);
            securableObject.AddPermissionLevelImplementation(principal, roleDefinition, removeExistingPermissionLevels);
        }

        private static void AddPermissionLevelImplementation(this SecurableObject securableObject, Principal principal, RoleDefinition roleDefinition, bool removeExistingPermissionLevels = false)
        {
            if (principal == null)
            {
                return;
            }

            var roleAssignments = securableObject.RoleAssignments;
            securableObject.Context.Load(roleAssignments);
            securableObject.Context.ExecuteQueryRetry();

            var roleAssignment = roleAssignments.FirstOrDefault(ra => ra.PrincipalId.Equals(principal.Id));

            //current principal doesn't have any roles assigned for this securableObject
            if (roleAssignment == null)
            {
                var rdc = new RoleDefinitionBindingCollection(securableObject.Context);
                rdc.Add(roleDefinition);
                securableObject.RoleAssignments.Add(principal, rdc);
                securableObject.Context.ExecuteQueryRetry();
            }
            else //current principal has roles assigned for this securableObject, then add new role definition for the role assignment
            {
                var roleDefinitionBindings = roleAssignment.RoleDefinitionBindings;
                securableObject.Context.Load(roleDefinitionBindings);
                securableObject.Context.ExecuteQueryRetry();

                // Load the role definition to add (e.g. contribute)
                if (removeExistingPermissionLevels)
                {
                    // Remove current role definitions by removing all current role definitions
                    roleDefinitionBindings.RemoveAll();
                }
                // Add the selected role definition
                if (!roleDefinitionBindings.Any(r => r.Name.Equals(roleDefinition.EnsureProperty(rd => rd.Name))))
                {
                    roleDefinitionBindings.Add(roleDefinition);

                    //update
                    roleAssignment.ImportRoleDefinitionBindings(roleDefinitionBindings);
                    roleAssignment.Update();
                    securableObject.Context.ExecuteQueryRetry();
                }
            }
        }

        /// <summary>
        /// Removes a permission level from a user
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="userLoginName">Loginname of user</param>
        /// <param name="permissionLevel">Permission level to remove. If null all permission levels are removed</param>
        /// <param name="removeAllPermissionLevels">Set to true to remove all permission level.</param>
        public static void RemovePermissionLevelFromUser(this SecurableObject securableObject, string userLoginName, RoleType permissionLevel, bool removeAllPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            Web web = securableObject.GetAssociatedWeb();

            User user = web.EnsureUser(userLoginName);

            securableObject.RemovePermissionLevelFromPrincipal(user, permissionLevel, removeAllPermissionLevels);
        }

        /// <summary>
        /// Removes a permission level from a user
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="principal">Principal to remove permission from</param>
        /// <param name="permissionLevel">Permission level to remove. If null all permission levels are removed</param>
        /// <param name="removeAllPermissionLevels">Set to true to remove all permission level.</param>
        public static void RemovePermissionLevelFromPrincipal(this SecurableObject securableObject, Principal principal, RoleType permissionLevel, bool removeAllPermissionLevels = false)
        {
            if (principal == null)
                throw new ArgumentNullException("principal");

            Web web = securableObject.GetAssociatedWeb();

            securableObject.Context.Load(principal);
            securableObject.Context.ExecuteQueryRetry();
            RoleDefinition roleDefinition = web.RoleDefinitions.GetByType(permissionLevel);

            securableObject.RemovePermissionLevelImplementation(principal, roleDefinition, removeAllPermissionLevels);
        }

        /// <summary>
        /// Removes a permission level from a user
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="userLoginName">Loginname of user</param>
        /// <param name="roleDefinitionName">Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Heirarchy|Restricted Read. Use the correct name of the language of the site you are using</param>
        /// <param name="removeAllPermissionLevels">Set to true to remove all permission level.</param>
        public static void RemovePermissionLevelFromUser(this SecurableObject securableObject, string userLoginName, string roleDefinitionName, bool removeAllPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            Web web = securableObject.GetAssociatedWeb();

            User user = web.EnsureUser(userLoginName);

            securableObject.RemovePermissionLevelFromPrincipal(user, roleDefinitionName, removeAllPermissionLevels);
        }

        /// <summary>
        /// Removes a permission level from a user
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="principal">Principal to remove permission from</param>
        /// <param name="roleDefinitionName">Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Heirarchy|Restricted Read. Use the correct name of the language of the site you are using</param>
        /// <param name="removeAllPermissionLevels">Set to true to remove all permission level.</param>
        public static void RemovePermissionLevelFromPrincipal(this SecurableObject securableObject, Principal principal, string roleDefinitionName, bool removeAllPermissionLevels = false)
        {
            if (principal == null)
                throw new ArgumentNullException("principal");

            Web web = securableObject.GetAssociatedWeb();

            securableObject.Context.Load(principal);
            securableObject.Context.ExecuteQueryRetry();
            RoleDefinition roleDefinition = web.RoleDefinitions.GetByName(roleDefinitionName);

            securableObject.RemovePermissionLevelImplementation(principal, roleDefinition, removeAllPermissionLevels);
        }

        /// <summary>
        /// Removes a permission level from a group
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="groupName">name of the group</param>
        /// <param name="permissionLevel">Permission level to remove. If null all permission levels are removed</param>
        /// <param name="removeAllPermissionLevels">Set to true to remove all permission level.</param>
        public static void RemovePermissionLevelFromGroup(this SecurableObject securableObject, string groupName, RoleType permissionLevel, bool removeAllPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            Web web = securableObject.GetAssociatedWeb();

            var group = web.SiteGroups.GetByName(groupName);
            securableObject.Context.Load(group);
            securableObject.Context.ExecuteQueryRetry();
            RoleDefinition roleDefinition = web.RoleDefinitions.GetByType(permissionLevel);
            securableObject.RemovePermissionLevelImplementation(group, roleDefinition, removeAllPermissionLevels);
        }

        /// <summary>
        /// Removes a permission level from a group
        /// </summary>
        /// <param name="securableObject">Web/List/Item to operate against</param>
        /// <param name="groupName">name of the group</param>
        /// <param name="roleDefinitionName">Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Heirarchy|Restricted Read. Use the correct name of the language of the site you are using</param>
        /// <param name="removeAllPermissionLevels">Set to true to remove all permission level.</param>
        public static void RemovePermissionLevelFromGroup(this SecurableObject securableObject, string groupName, string roleDefinitionName, bool removeAllPermissionLevels = false)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            Web web = securableObject.GetAssociatedWeb();

            var group = web.SiteGroups.GetByName(groupName);
            securableObject.Context.Load(group);
            securableObject.Context.ExecuteQueryRetry();
            RoleDefinition roleDefinition = web.RoleDefinitions.GetByName(roleDefinitionName);
            securableObject.RemovePermissionLevelImplementation(group, roleDefinition, removeAllPermissionLevels);
        }

        private static void RemovePermissionLevelImplementation(this SecurableObject securableObject, Principal principal, RoleDefinition roleDefinition, bool removeAllPermissionLevels = false)
        {
            if (principal == null)
            {
                return;
            }

            var roleAssignments = securableObject.RoleAssignments;
            securableObject.Context.Load(roleAssignments);
            securableObject.Context.ExecuteQueryRetry();

            var roleAssignment = roleAssignments.FirstOrDefault(ra => ra.PrincipalId.Equals(principal.Id));

            if (roleAssignment != null)
            {
                // load the role definitions for this role assignment
                var rdc = roleAssignment.RoleDefinitionBindings;
                securableObject.Context.Load(rdc);
                securableObject.Context.ExecuteQueryRetry();

                if (removeAllPermissionLevels)
                {
                    // Remove current role definitions by removing all current role definitions
                    rdc.RemoveAll();
                }
                else
                {
                    // Load the role definition to remove (e.g. contribute)
                    rdc.Remove(roleDefinition);
                }

                //update
                roleAssignment.ImportRoleDefinitionBindings(rdc);
                roleAssignment.Update();
                securableObject.Context.ExecuteQueryRetry();
            }
        }

        private static Web GetAssociatedWeb(this SecurableObject securable)
        {
            if (securable is Web)
            {
                return (Web)securable;
            }

            if (securable is List)
            {
                var list = (List)securable;
                var web = list.ParentWeb;
                securable.Context.Load(web);
                securable.Context.ExecuteQueryRetry();

                return web;
            }

            if (securable is ListItem)
            {
                var listItem = (ListItem)securable;
                var web = listItem.ParentList.ParentWeb;
                securable.Context.Load(web);
                securable.Context.ExecuteQueryRetry();

                return web;
            }

            throw new Exception("Only Web, List, ListItem supported as SecurableObjects");
        }

        /// <summary>
        /// Removes a user from a group
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="groupName">Name of the group</param>
        /// <param name="userLoginName">Loginname of the user</param>
        public static void RemoveUserFromGroup(this Web web, string groupName, string userLoginName)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            var group = web.SiteGroups.GetByName(groupName);
            web.Context.Load(group);
            web.Context.ExecuteQueryRetry();
            if (group != null)
            {
                // Check wether the group contains any users to avoid errors / exceptions
                group.EnsureProperty(g => g.Users);
                if (group.Users.Count > 0)
                {
                    User user = group.Users.GetByLoginName(userLoginName);
                    web.Context.Load(user);
                    web.Context.ExecuteQueryRetry();
                    if (!user.ServerObjectIsNull.Value)
                    {
                        web.RemoveUserFromGroup(group, user);
                    }
                }
            }
        }

        /// <summary>
        /// Removes a user from a group
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="group">Group object to operate against</param>
        /// <param name="user">User object that needs to be removed</param>
        public static void RemoveUserFromGroup(this Web web, Group group, User user)
        {
            if (group == null)
                throw new ArgumentNullException("group");

            if (user == null)
                throw new ArgumentNullException("user");

            group.Users.Remove(user);
            group.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Remove a group
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="groupName">Name of the group</param>
        public static void RemoveGroup(this Web web, string groupName)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            var group = web.SiteGroups.GetByName(groupName);
            web.Context.Load(group);
            web.Context.ExecuteQueryRetry();
            if (group != null)
            {
                web.RemoveGroup(group);
            }
        }

        /// <summary>
        /// Remove a group
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="group">Group object to remove</param>
        public static void RemoveGroup(this Web web, Group group)
        {
            if (group == null)
                throw new ArgumentNullException("group");

            GroupCollection groups = web.SiteGroups;
            groups.Remove(group);
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Checks if a user is member of a group
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="groupName">Name of the group</param>
        /// <param name="userLoginName">Loginname of the user</param>
        /// <returns>True if the user is in the group, false otherwise</returns>
        public static bool IsUserInGroup(this Web web, string groupName, string userLoginName)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            if (string.IsNullOrEmpty(userLoginName))
                throw new ArgumentNullException("userLoginName");

            bool result = false;

            var group = web.SiteGroups.GetByName(groupName);
            var users = group.Users;
            web.Context.Load(group);
            web.Context.Load(users);
            web.Context.ExecuteQueryRetry();

            if (users.AreItemsAvailable)
            {
                result = users.Any(u =>
                  u.LoginName.ToLowerInvariant().Contains(userLoginName.ToLowerInvariant())
                );
            }

            return result;
        }

        /// <summary>
        /// Checks if a group exists
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="groupName">Name of the group</param>
        /// <returns>True if the group exists, false otherwise</returns>
        public static bool GroupExists(this Web web, string groupName)
        {
            if (string.IsNullOrEmpty(groupName))
                throw new ArgumentNullException("groupName");

            bool result = false;

            try
            {
                var group = web.SiteGroups.GetByName(groupName);
                web.Context.Load(group);
                web.Context.ExecuteQueryRetry();
                if (group != null)
                {
                    result = true;
                }
            }
            catch (ServerException ex)
            {
                if (IsGroupCannotBeFoundException(ex))
                {
                    //eat the exception
                }
                else
                {
                    //rethrow exception
                    throw;
                }
            }

            return result;
        }

        private static bool IsGroupCannotBeFoundException(Exception ex)
        {
            if (ex is ServerException)
            {
                if (((ServerException)ex).ServerErrorCode == -2146232832 && ((ServerException)ex).ServerErrorTypeName.Equals("Microsoft.SharePoint.SPException", StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        #endregion


        #region Authentication Realm
        /// <summary>
        /// Returns the authentication realm for the current web
        /// </summary>
        /// <param name="web">The Current site</param>
        /// <returns>Returns Realm in Guid</returns>
        public static Guid GetAuthenticationRealm(this Web web)
        {
            web.EnsureProperty(w => w.Url);
#if !NETSTANDARD2_0
            Guid.TryParse(TokenHelper.GetRealmFromTargetUrl(new Uri(web.Url)), out var g);            
            return g;
#else
            WebRequest request = WebRequest.Create(new Uri(web.Url) + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                var bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];

                const string bearer = "Bearer realm=\"";
                var bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);

                var realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    var targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return realmGuid;
                    }
                }
            }
            return Guid.Empty;
#endif
        }
        #endregion


        #region SecurableObject traversal

        #region Helpers

        /// <summary>
        /// Get URL path of a securable object
        /// </summary>
        /// <param name="obj">A securable object which could be a web, a list, a list item, a document library or a document</param>
        /// <returns>The URL of the securable object</returns>
        internal static string GetPath(this SecurableObject obj)
        {
            var path = String.Empty;
            if (obj is Web)
            {
                var web = obj as Web;
                path = web.Url;
            }
            else if (obj is List)
            {
                var list = obj as List;
                path = list.DefaultViewUrl;
                var defaultPages = Constants.DefaultViewPages.Where(p => path.EndsWith(p));
                if (defaultPages.Any())
                    path = path.Substring(0, path.Length - defaultPages.First().Length);
            }
            else if (obj is ListItem)
            {
                var item = obj as ListItem;
                path = string.Format("{0}/{1}", item.FieldValues[Constants.ListItemDirField], item.FieldValues[Constants.ListItemFileNameField]);
            }
            return path;
        }

        /// <summary>
        /// Load properties of the current securable object and get child securable objects with unique role assignments if any.
        /// </summary>
        /// <param name="obj">The current securable object.</param>
        /// <param name="leafBreadthLimit">Skip further visiting on this branch if the number of child items or documents with unique role assignments exceeded leafBreadthLimit.</param>
        /// <returns>The child securable objects.</returns>
        internal static IEnumerable<SecurableObject> Preload(this SecurableObject obj, int leafBreadthLimit)
        {
            var context = obj.Context;
            IEnumerable<SecurableObject> subObjects = new SecurableObject[] { };
            if (obj is Web)
            {
                var web = obj as Web;
                context.Load(web, w => w.Url, w => w.HasUniqueRoleAssignments);
                context.ExecuteQueryRetry();
                if (web.HasUniqueRoleAssignments)
                {
                    context.Load(web.RoleAssignments);
                    context.Load(web.RoleAssignments.Groups);
                    context.ExecuteQueryRetry();
                }
                var lists = context.LoadQuery(web.Lists.Where(l => l.BaseType == BaseType.DocumentLibrary));
                var webs = context.LoadQuery(web.Webs);
                context.ExecuteQueryRetry();
                subObjects = lists.Select(l => l as SecurableObject).Union(webs.Select(w => w as SecurableObject));
            }
            else if (obj is List)
            {
                var list = obj as List;
                context.Load(list, l => l.ItemCount, l => l.DefaultViewUrl, l => l.HasUniqueRoleAssignments);
                context.ExecuteQueryRetry();
                if (list.ItemCount <= 0 || Constants.SkipPathes.Any(p => list.DefaultViewUrl.IndexOf(p) != -1))
                    return null;
                if (list.HasUniqueRoleAssignments)
                {
                    context.Load(list.RoleAssignments);
                    context.Load(list.RoleAssignments.Groups);
                    context.ExecuteQueryRetry();
                }
                if (leafBreadthLimit > 0 && list.ItemCount > 0)
                {
                    var query = new CamlQuery();
                    query.ViewXml = String.Format(Constants.AllItemCamlQuery, Constants.ListItemDirField, Constants.ListItemFileNameField);
                    var items = context.LoadQuery(list.GetItems(query).Where(i => i.HasUniqueRoleAssignments));
                    context.ExecuteQueryRetry();
                    if (items.Count() <= leafBreadthLimit)
                    {
                        subObjects = items;
                    }
                    else
                    {
                        Trace.TraceWarning(CoreResources.SecurityExtensions_Warning_SkipFurtherVisitingForTooManyChildObjects, obj.GetPath(), list.ItemCount, leafBreadthLimit);
                    }
                }
            }
            else if (obj is ListItem)
            {
                var item = obj as ListItem;
                context.Load(item, i => i.HasUniqueRoleAssignments);
                context.Load(item.RoleAssignments);
                context.Load(item.RoleAssignments.Groups);
                context.ExecuteQueryRetry();
            }
            return subObjects;
        }

        /// <summary>
        /// Traverse each descendents of a securable object with a specified action.
        /// </summary>
        /// <param name="obj">The current securable object.</param>
        /// <param name="leafBreadthLimit">Skip further visiting on this branch if the number of child items or documents with unique role assignments exceeded leafBreadthLimit.</param>
        /// <param name="action">The action to be executed for each securable object.</param>
        internal static void Visit(this SecurableObject obj, int leafBreadthLimit, Action<SecurableObject, string> action)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");
            if (action == null)
                throw new ArgumentNullException("action");
            string path = string.Empty;
            var stack = new Stack<SecurableObject>();
            stack.Push(obj);
            while (stack.Count != 0)
            {
                try
                {
                    var currentObj = stack.Pop();
                    var subObjects = currentObj.Preload(leafBreadthLimit);
                    if (subObjects == null)
                        continue;
                    path = currentObj.GetPath();
                    Trace.TraceInformation(CoreResources.SecurityExtensions_Info_VisitingSecurableObject, path);
                    action(currentObj, path);
                    foreach (var subObj in subObjects.Reverse())
                    {
                        stack.Push(subObj);
                    }
                }
                catch (Exception e)
                {
                    Trace.TraceError(CoreResources.SecurityExtensions_Error_VisitingSecurableObject, path, e);
                }
            }
        }

        #endregion

        #region Entity cache

        /// <summary>
        /// A dictionary to cache resolved user emails. key: user login name, value: user email.
        /// ***
        /// Don't use this cache in a real world application.
        /// ***
        /// Instead it should be replaced by a real cache with ref object to clear up intermediate records periodically.
        /// </summary>
        private static Dictionary<string, string> MockupUserEmailCache = new Dictionary<string, string>();

        /// <summary>
        /// Get user email by user id.
        /// </summary>
        /// <param name="web">The current web object.</param>
        /// <param name="userId">The user id</param>
        /// <returns>The email property of the specified user.</returns>
        private static string GetUserEmail(this Web web, int userId)
        {
            var user = web.GetUserById(userId);
            web.Context.Load(user, u => u.Email);
            web.Context.ExecuteQueryRetry();
            return user.Email;
        }

        /// <summary>
        /// A dictionary to cache all user entities of a given SharePoint group. key: group login name, value: an array of user entities belongs to the group.
        /// ***
        /// Don't use this cache in a real world application.
        /// ***
        /// Instead it should be replaced by a real cache with ref object to clear up intermediate records periodically.
        /// </summary>
        private static Dictionary<string, UserEntity[]> MockupGroupCache = new Dictionary<string, UserEntity[]>();

        /// <summary>
        /// Ensure all users of a given SharePoint group has been cached.
        /// </summary>
        /// <param name="obj">The current securable object.</param>
        /// <param name="groupLoginName">The group login name.</param>
        private static void EnsureGroupCache(SecurableObject obj, string groupLoginName)
        {
            var context = obj.Context;
            if (!MockupGroupCache.ContainsKey(groupLoginName))
            {
                var users = context.LoadQuery(obj.RoleAssignments.Groups.First(g => g.LoginName.Equals(groupLoginName, StringComparison.OrdinalIgnoreCase)).Users);
                context.ExecuteQueryRetry();
                MockupGroupCache[groupLoginName] = (from u in users select new UserEntity()
                                                    {
                                                        Title = u.Title,
                                                        Email = u.Email,
                                                        LoginName = u.LoginName
                                                    }).ToArray();
            }
        }

        #endregion

        /// <summary>
        /// Get all unique role assignments for a web object and all its descendents down to document or list item level.
        /// </summary>
        /// <param name="web">The current web object to be processed.</param>
        /// <param name="leafBreadthLimit">Skip further visiting on this branch if the number of child items or documents with unique role assignments exceeded leafBreadthLimit. When setting to 0, the process will stop at list / document library level.</param>
        /// <returns>Returns all role assignments</returns>
        public static IEnumerable<RoleAssignmentEntity> GetAllUniqueRoleAssignments(this Web web, int leafBreadthLimit = int.MaxValue)
        {
            var result = new List<RoleAssignmentEntity>();
            web.Visit(leafBreadthLimit, (obj, path) =>
            {
                if (!obj.HasUniqueRoleAssignments)
                {
                    return;
                }
                foreach (var assignment in obj.RoleAssignments)
                {
                    var bindings = web.Context.LoadQuery(assignment.RoleDefinitionBindings.Where(b => b.Name != "Limited Access"));
                    web.Context.Load(assignment.Member, m => m.LoginName, m => m.Title, m => m.PrincipalType, m => m.Id);
                    web.Context.ExecuteQueryRetry();
                    var bindingList = (from b in bindings select b.Name).ToList();
                    if (assignment.Member.PrincipalType == Utilities.PrincipalType.SharePointGroup)
                    {
                        EnsureGroupCache(obj, assignment.Member.LoginName);
                        foreach (var user in MockupGroupCache[assignment.Member.LoginName])
                        {
                            if (!MockupUserEmailCache.ContainsKey(user.LoginName))
                            {
                                MockupUserEmailCache[user.LoginName] = user.Email;
                            }
                            result.Add(new RoleAssignmentEntity()
                            {
                                Path = path,
                                User = user,
                                Role = assignment.Member.Title,
                                RoleDefinitionBindings = bindingList,
                            });
                        }
                    }
                    else
                    {
                        if (!MockupUserEmailCache.ContainsKey(assignment.Member.LoginName))
                        {
                            MockupUserEmailCache[assignment.Member.LoginName] = web.GetUserEmail(assignment.Member.Id);
                        }
                        result.Add(new RoleAssignmentEntity()
                        {
                            Path = path,
                            User = new UserEntity()
                            {
                                Title = assignment.Member.Title,
                                Email = MockupUserEmailCache[assignment.Member.LoginName],
                                LoginName = assignment.Member.LoginName
                            },
                            Role = "(Directly Assigned)",
                            RoleDefinitionBindings = bindingList,
                        });
                    }
                }
            });
            return result;
        }

        /// <summary>
        /// Get all unique role assignments for a user or a group in a web object and all its descendents down to document or list item level.
        /// </summary>
        /// <param name="web">The current web object to be processed.</param>
        /// <param name="principal">The current web object to be processed.</param>
        /// <param name="leafBreadthLimit">Skip further visiting on this branch if the number of child items or documents with unique role assignments exceeded leafBreadthLimit. When setting to 0, the process will stop at list / document library level.</param>
        /// <returns>Returns all role assignments</returns>
        public static IEnumerable<RoleAssignmentEntity> GetPrincipalUniqueRoleAssignments(this Web web, Principal principal, int leafBreadthLimit = int.MaxValue)
        {
            var result = new List<RoleAssignmentEntity>();
            web.Visit(leafBreadthLimit, (obj, path) =>
            {
                if (!obj.HasUniqueRoleAssignments)
                {
                    return;
                }

                RoleAssignment assignment;
                try
                {
                    assignment = obj.RoleAssignments.GetByPrincipal(principal);
                    web.Context.ExecuteQueryRetry();
                }
                catch (ServerException ex)
                {
                    if (ex.HResult == -2146233088)
                    {
                        // Can not find the principal so suppress the exception
                        assignment = null;
                    }
                    else
                    {
                        throw;
                    }
                }

                if (assignment != null)
                {
                    var bindings = web.Context.LoadQuery(assignment.RoleDefinitionBindings.Where(b => b.Name != "Limited Access"));
                    web.Context.Load(assignment.Member, m => m.LoginName, m => m.Title, m => m.PrincipalType, m => m.Id);
                    web.Context.ExecuteQueryRetry();

                    var bindingList = (from b in bindings select b.Name).ToList();
                    if (bindingList.Count != 0)
                    {
                        if (assignment.Member.PrincipalType == Utilities.PrincipalType.SharePointGroup)
                        {
                            EnsureGroupCache(obj, assignment.Member.LoginName);
                            foreach (var user in MockupGroupCache[assignment.Member.LoginName])
                            {
                                if (!MockupUserEmailCache.ContainsKey(user.LoginName))
                                {
                                    MockupUserEmailCache[user.LoginName] = user.Email;
                                }
                                result.Add(new RoleAssignmentEntity()
                                {
                                    Path = path,
                                    User = user,
                                    Role = assignment.Member.Title,
                                    RoleDefinitionBindings = bindingList,
                                });
                            }
                        }
                        else
                        {
                            if (!MockupUserEmailCache.ContainsKey(assignment.Member.LoginName))
                            {
                                MockupUserEmailCache[assignment.Member.LoginName] = web.GetUserEmail(assignment.Member.Id);
                            }
                            result.Add(new RoleAssignmentEntity()
                            {
                                Path = path,
                                User = new UserEntity()
                                {
                                    Title = assignment.Member.Title,
                                    Email = MockupUserEmailCache[assignment.Member.LoginName],
                                    LoginName = assignment.Member.LoginName
                                },
                                Role = "(Directly Assigned)",
                                RoleDefinitionBindings = bindingList,
                            });
                        }
                    }
                }
            });
            return result;
        }
        #endregion
    }
}
