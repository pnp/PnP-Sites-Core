using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// Provides function for checking permission level of user/sharepoint group
    /// </summary>
    public static class CheckPermissionUtility
    {
        /// <summary>
        /// This function returns
        /// collection of permission level
        /// of an user/group 
        /// for a web/list/list item
        /// </summary>
        /// <param name="securableObject">web/list/list item</param>
        /// <param name="context">ClientContext</param>
        /// <param name="principalName">user (email/login)/group name</param>
        /// <returns></returns>
        public static List<string> GetPrincipalsRoleDefinitions(this SecurableObject securableObject, ClientContext context, string principalName)
        {
            try
            {
                var permissionDefinitions = new List<string>();

                var principal = ResolvePrincipal(context, principalName);
                if (principal == null) throw new Exception("Colud not resolve Principal.");

                var roleAssignments = GetRoleAssignments(context, securableObject);

                var directPermissionLevel = roleAssignments.Where(r => r.Key == principal.Title).FirstOrDefault();
                permissionDefinitions.AddRange(directPermissionLevel.Value.TrimEnd(',').Split(','));

                if (principal.PrincipalType == PrincipalType.User)
                {
                    var user = context.Web.EnsureUser(principalName);
                    context.Load(user, u => u.LoginName);
                    context.ExecuteQueryRetry();

                    var groups = GetUserGroups(context, user.LoginName);

                    var userGroupsPermissionLevel = (from g in groups
                                                     join r in roleAssignments
                                                     on g equals r.Key
                                                     select new
                                                     {
                                                         GroupName = g,
                                                         PermissionLevel = r.Value
                                                     });

                    foreach (var r in userGroupsPermissionLevel)
                        permissionDefinitions.AddRange(r.PermissionLevel.TrimEnd(',').Split(','));
                }

                return permissionDefinitions.Distinct().ToList();
            }
            catch (Exception ex)
            {
                Diagnostics.Log.Error(Constants.LOGGING_SOURCE, "", ex);
                throw;
            }
        }

        private static List<string> GetUserGroups(ClientContext context, string userLogin)
        {
            User user = context.Web.SiteUsers.GetByLoginName(userLogin);
            GroupCollection groupColl = user.Groups;

            context.Load(groupColl);
            context.ExecuteQueryRetry();

            return groupColl.Select(g => g.Title).ToList();
        }

        private static Dictionary<string, string> GetRoleAssignments(ClientContext context, SecurableObject securableObject)
        {
            IQueryable<RoleAssignment> query = securableObject.RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                                   roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));

            Dictionary<string, string> assignments = GetRoleDefinitionBindings(context, query);

            return assignments;
        }

        private static Dictionary<string, string> GetRoleDefinitionBindings(ClientContext context, IQueryable<RoleAssignment> queryString)
        {
            IEnumerable roles = context.LoadQuery(queryString);
            context.ExecuteQueryRetry();

            Dictionary<string, string> permisionDetails = new Dictionary<string, string>();
            foreach (RoleAssignment ra in roles)
            {
                var rdc = ra.RoleDefinitionBindings;
                string permissionLevel = string.Empty;

                foreach (var rdbc in rdc)
                {
                    permissionLevel += rdbc.Name.ToString() + ",";
                }

                permisionDetails.Add(ra.Member.Title, permissionLevel);
            }

            return permisionDetails;
        }

        private static Principal ResolvePrincipal(ClientContext context, string name)
        {
            Principal principal = null;

            try
            {
                ClientResult<PrincipalInfo> result = Microsoft.SharePoint.Client.Utilities.Utility.ResolvePrincipal(context, context.Web, name, PrincipalType.All, PrincipalSource.All, null, false);
                context.ExecuteQuery();

                if (result.Value.PrincipalType == PrincipalType.User)
                {
                    principal = context.Web.EnsureUser(result.Value.LoginName);
                }
                else if (result.Value.PrincipalType == PrincipalType.SecurityGroup || result.Value.PrincipalType == PrincipalType.SharePointGroup)
                {
                    if (result.Value.DisplayName == "")  // invalid input
                    {
                        return principal;
                    }
                    else
                    {
                        // sharepoint group -> principal type: Security Group, principal id: -1, login name: name of the group, email: null, display name: name of the group
                        if (result.Value.PrincipalId != -1)
                        {
                            principal = context.Web.SiteGroups.GetById(result.Value.PrincipalId);
                        }
                        // distribution list -> principal type: Security Group, principal id: -1, login name: c:0t.c|tenant|GUID_of_distribution_list, email: email address of the distribution list, display name: name of the disribution list
                        // special group -> principal type: Security Group, principal id: -1, login name: c:0-.f|rolemanager|spo-grid-all-users/GUID, email: null, display name: Everyone except external users
                        else
                        {
                            principal = context.Web.EnsureUser(result.Value.LoginName);
                        }
                    }
                }

                context.Load(principal);
                context.ExecuteQueryRetry();
            }
            catch (Exception)
            {
                principal = null;
            }

            return principal;
        }
    }
}