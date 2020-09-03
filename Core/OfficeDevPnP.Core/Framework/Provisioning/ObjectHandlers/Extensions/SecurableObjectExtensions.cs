using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{
    internal static class SecurableObjectExtensions
    {
        private static Principal TryGetGroupPrincipal(IEnumerable<Microsoft.SharePoint.Client.Group> groups, string roleAssignmentPrincipal)
        {
            Principal principal = groups.FirstOrDefault(g => g.LoginName.Equals(roleAssignmentPrincipal, StringComparison.OrdinalIgnoreCase));

            // Principal can be resolved via it's ID if an associatedgroupid token was used
            if (principal == null)
            {
                if (Int32.TryParse(roleAssignmentPrincipal, out int roleAssignmentPrincipalId))
                {
                    principal = groups.FirstOrDefault(g => g.Id.Equals(roleAssignmentPrincipalId));
                }
            }

            return principal;
        }

        private static IEnumerable<Model.RoleAssignment> CheckForAndRemoveNonExistingPrincipals(IEnumerable<Model.RoleAssignment> RoleAssignments, TokenParser parser, IEnumerable<Microsoft.SharePoint.Client.Group> groups, ClientContext context, ProvisioningMessagesDelegate MessageDelegate)
        {
            var result = new List<Model.RoleAssignment>();
            foreach (var roleAssignment in RoleAssignments)
            {
                var roleAssignmentPrincipal = parser.ParseString(roleAssignment.Principal);

                Principal principal = TryGetGroupPrincipal(groups, roleAssignmentPrincipal);

                if (principal == null)
                {
                    try
                    {
                        context.Web.EnsureUser(roleAssignmentPrincipal);
                        context.ExecuteQueryRetry();
                    }
                    catch (ServerException ex)
                    {
                        // catch user not found
                        if (ex.ServerErrorCode == -2146232832 && ex.ServerErrorTypeName.Equals("Microsoft.SharePoint.SPException", StringComparison.InvariantCultureIgnoreCase))
                        {
                            MessageDelegate($"Cannot find principal '{roleAssignmentPrincipal}', cannot grant permissions", ProvisioningMessageType.Warning);
                            continue;
                        }
                    }
                }

                result.Add(roleAssignment);
            }
            return result;
        }

        private static void ApplySecurity(SecurableObject securable, TokenParser parser, ClientContext context, IEnumerable<Microsoft.SharePoint.Client.Group> groups, IEnumerable<Microsoft.SharePoint.Client.RoleDefinition> webRoleDefinitions, IEnumerable<Microsoft.SharePoint.Client.RoleAssignment> securableRoleAssignments, IEnumerable<Model.RoleAssignment> roleAssignmentsToHandle)
        {
            foreach (var roleAssignment in roleAssignmentsToHandle)
            {
                if (!roleAssignment.Remove)
                {
                    var roleAssignmentPrincipal = parser.ParseString(roleAssignment.Principal);

                    Principal principal = TryGetGroupPrincipal(groups, roleAssignmentPrincipal);

                    if (principal == null)
                    {
                        principal = context.Web.EnsureUser(roleAssignmentPrincipal);
                    }

                    if (principal != null)
                    {
                        var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(context);

                        var roleAssignmentRoleDefinition = parser.ParseString(roleAssignment.RoleDefinition);
                        var roleDefinition = webRoleDefinitions.FirstOrDefault(r => r.Name == roleAssignmentRoleDefinition);

                        if (roleDefinition != null)
                        {
                            roleDefinitionBindingCollection.Add(roleDefinition);
                            securable.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
                        }
                    }
                }
                else
                {
                    var roleAssignmentPrincipal = parser.ParseString(roleAssignment.Principal);

                    Principal principal = TryGetGroupPrincipal(groups, roleAssignmentPrincipal);

                    if (principal == null)
                    {
                        principal = context.Web.EnsureUser(roleAssignmentPrincipal);
                    }
                    principal.EnsureProperty(p => p.Id);

                    if (principal != null)
                    {
                        var assignmentsForPrincipal = securableRoleAssignments.Where(t => t.PrincipalId == principal.Id);
                        foreach (var assignmentForPrincipal in assignmentsForPrincipal)
                        {
                            var roleAssignmentRoleDefinition = parser.ParseString(roleAssignment.RoleDefinition);
                            var binding = assignmentForPrincipal.EnsureProperty(r => r.RoleDefinitionBindings).FirstOrDefault(b => b.Name == roleAssignmentRoleDefinition);
                            if (binding != null)
                            {
                                assignmentForPrincipal.DeleteObject();
                                context.ExecuteQueryRetry();
                                break;
                            }
                        }
                    }
                }
            }
        }

        public static void SetSecurity(this SecurableObject securable, TokenParser parser, ObjectSecurity security, ProvisioningMessagesDelegate MessageDelegate)
        {
            // If there's no role assignments we're returning
            if (security.RoleAssignments.Count == 0) return;

            var context = securable.Context as ClientContext;

            var groups = context.LoadQuery(context.Web.SiteGroups.Include(g => g.LoginName, g => g.Id));
            var webRoleDefinitions = context.LoadQuery(context.Web.RoleDefinitions);

            securable.BreakRoleInheritance(security.CopyRoleAssignments, security.ClearSubscopes);

            var securableRoleAssignments = context.LoadQuery(securable.RoleAssignments);
            context.ExecuteQueryRetry();
            IEnumerable<Model.RoleAssignment> roleAssignmentsToHandle = security.RoleAssignments;

            // try to apply the security in two steps: step one assumes all principals from the template exist and can be granted permission at once
            try
            {
                // note that this step fails if there is one principal that doesn't exist
                ApplySecurity(securable, parser, context, groups, webRoleDefinitions, securableRoleAssignments, roleAssignmentsToHandle);
                context.ExecuteQueryRetry();
            }
            catch (ServerException ex)
            {
                // catch user not found; enter step 2: check each and every principal for existence before granting security for those that exist
                if (ex.ServerErrorCode == -2146232832 && ex.ServerErrorTypeName.Equals("Microsoft.SharePoint.SPException", StringComparison.InvariantCultureIgnoreCase))
                {
                    roleAssignmentsToHandle = CheckForAndRemoveNonExistingPrincipals(roleAssignmentsToHandle, parser, groups, context, MessageDelegate);
                    ApplySecurity(securable, parser, context, groups, webRoleDefinitions, securableRoleAssignments, roleAssignmentsToHandle);
                    // if it fails this time we just let it fail
                    context.ExecuteQueryRetry();
                }
            }
        }

        public static ObjectSecurity GetSecurity(this SecurableObject securable)
        {
            ObjectSecurity security = null;
            using (var scope = new PnPMonitoredScope("Get Security"))
            {
                var context = securable.Context as ClientContext;

                context.Load(securable, sec => sec.HasUniqueRoleAssignments);
                context.Load(context.Web, w => w.AssociatedMemberGroup.Title, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedVisitorGroup.Title);
                var roleAssignments = context.LoadQuery(securable.RoleAssignments.Include(
                    r => r.Member.LoginName,
                    r => r.RoleDefinitionBindings.Include(
                        rdb => rdb.Name,
                        rdb => rdb.RoleTypeKind
                        )));
                context.ExecuteQueryRetry();

                if (securable.HasUniqueRoleAssignments)
                {
                    security = new ObjectSecurity();

                    foreach (var roleAssignment in roleAssignments)
                    {
                        if (roleAssignment.Member.LoginName != "Excel Services Viewers")
                        {
                            foreach (var roleDefinition in roleAssignment.RoleDefinitionBindings)
                            {
                                if (roleDefinition.RoleTypeKind != RoleType.Guest)
                                {
                                    security.RoleAssignments.Add(new Model.RoleAssignment()
                                    {
                                        Principal = ReplaceGroupTokens(context.Web, roleAssignment.Member.LoginName),
                                        RoleDefinition = roleDefinition.Name
                                    });
                                }
                            }
                        }
                    }
                }
            }
            return security;
        }

        private static string ReplaceGroupTokens(Web web, string loginName)
        {
            if (web.AssociatedOwnerGroup.ServerObjectIsNull.HasValue && !web.AssociatedOwnerGroup.ServerObjectIsNull.Value)
            {
                loginName = loginName.Replace(web.AssociatedOwnerGroup.Title, "{associatedownergroupid}");
            }
            if (web.AssociatedMemberGroup.ServerObjectIsNull.HasValue && !web.AssociatedMemberGroup.ServerObjectIsNull.Value)
            {
                loginName = loginName.Replace(web.AssociatedMemberGroup.Title, "{associatedmembergroupid}");
            }
            if (web.AssociatedVisitorGroup.ServerObjectIsNull.HasValue && !web.AssociatedVisitorGroup.ServerObjectIsNull.Value)
            {
                loginName = loginName.Replace(web.AssociatedVisitorGroup.Title, "{associatedvisitorgroupid}");
            }
            return loginName;
        }
    }
}
