using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using User = OfficeDevPnP.Core.Framework.Provisioning.Model.User;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSiteSecurity : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Site Security"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {


                // if this is a sub site then we're not provisioning security as by default security is inherited from the root site
                if (web.IsSubSite())
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_SiteSecurity_Context_web_is_subweb__skipping_site_security_provisioning);
                    return parser;
                }

                var siteSecurity = template.Security;

                var ownerGroup = web.AssociatedOwnerGroup;
                var memberGroup = web.AssociatedMemberGroup;
                var visitorGroup = web.AssociatedVisitorGroup;


                web.Context.Load(ownerGroup, o => o.Title, o => o.Users);
                web.Context.Load(memberGroup, o => o.Title, o => o.Users);
                web.Context.Load(visitorGroup, o => o.Title, o => o.Users);

                web.Context.ExecuteQueryRetry();

                if (!ownerGroup.ServerObjectIsNull.Value)
                {
                    AddUserToGroup(web, ownerGroup, siteSecurity.AdditionalOwners, scope);
                }
                if (!memberGroup.ServerObjectIsNull.Value)
                {
                    AddUserToGroup(web, memberGroup, siteSecurity.AdditionalMembers, scope);
                }
                if (!visitorGroup.ServerObjectIsNull.Value)
                {
                    AddUserToGroup(web, visitorGroup, siteSecurity.AdditionalVisitors, scope);
                }

                foreach (var siteGroup in siteSecurity.SiteGroups)
                {
                    Group group = null;
                    var allGroups = web.Context.LoadQuery(web.SiteGroups.Include(gr => gr.LoginName));
                    web.Context.ExecuteQueryRetry();

                    if (!web.GroupExists(siteGroup.Title))
                    {
                        scope.LogDebug("Creating group {0}", siteGroup.Title);
                        group = web.AddGroup(
                            parser.ParseString(siteGroup.Title),
                            parser.ParseString(siteGroup.Description),
                            parser.ParseString(siteGroup.Title) == parser.ParseString(siteGroup.Owner));
                        group.AllowMembersEditMembership = siteGroup.AllowMembersEditMembership;
                        group.AllowRequestToJoinLeave = siteGroup.AllowRequestToJoinLeave;
                        group.AutoAcceptRequestToJoinLeave = siteGroup.AutoAcceptRequestToJoinLeave;

                        if (parser.ParseString(siteGroup.Title) != parser.ParseString(siteGroup.Owner))
                        {
                            Principal ownerPrincipal = allGroups.FirstOrDefault(gr => gr.LoginName == parser.ParseString(siteGroup.Owner));
                            if (ownerPrincipal == null)
                            {
                                ownerPrincipal = web.EnsureUser(parser.ParseString(siteGroup.Owner));
                            }
                            group.Owner = ownerPrincipal;

                        }
                        group.Update();
                        web.Context.ExecuteQueryRetry();
                    }
                    else
                    {
                        group = web.SiteGroups.GetByName(parser.ParseString(siteGroup.Title));
                        web.Context.Load(group,
                            g => g.Title,
                            g => g.Description,
                            g => g.AllowMembersEditMembership,
                            g => g.AllowRequestToJoinLeave,
                            g => g.AutoAcceptRequestToJoinLeave,
                            g => g.Owner.LoginName);
                        web.Context.ExecuteQueryRetry();
                        var isDirty = false;
                        if (group.Description != parser.ParseString(siteGroup.Description))
                        {
                            group.Description = parser.ParseString(siteGroup.Description);
                            isDirty = true;
                        }
                        if (group.AllowMembersEditMembership != siteGroup.AllowMembersEditMembership)
                        {
                            group.AllowMembersEditMembership = siteGroup.AllowMembersEditMembership;
                            isDirty = true;
                        }
                        if (group.AllowRequestToJoinLeave != siteGroup.AllowRequestToJoinLeave)
                        {
                            group.AllowRequestToJoinLeave = siteGroup.AllowRequestToJoinLeave;
                            isDirty = true;
                        }
                        if (group.AutoAcceptRequestToJoinLeave != siteGroup.AutoAcceptRequestToJoinLeave)
                        {
                            group.AutoAcceptRequestToJoinLeave = siteGroup.AutoAcceptRequestToJoinLeave;
                            isDirty = true;
                        }
                        if (group.Owner.LoginName != parser.ParseString(siteGroup.Owner))
                        {
                            if (parser.ParseString(siteGroup.Title) != parser.ParseString(siteGroup.Owner))
                            {
                                Principal ownerPrincipal = allGroups.FirstOrDefault(gr => gr.LoginName == parser.ParseString(siteGroup.Owner));
                                if (ownerPrincipal == null)
                                {
                                    ownerPrincipal = web.EnsureUser(parser.ParseString(siteGroup.Owner));
                                }
                                group.Owner = ownerPrincipal;
                            }
                            else
                            {
                                group.Owner = group;
                            }
                            isDirty = true;
                        }
                        if (isDirty)
                        {
                            scope.LogDebug("Updating existing group {0}", group.Title);
                            group.Update();
                            web.Context.ExecuteQueryRetry();
                        }
                    }
                    if (group != null && siteGroup.Members.Any())
                    {
                        AddUserToGroup(web, group, siteGroup.Members, scope);
                    }
                }

                foreach (var admin in siteSecurity.AdditionalAdministrators)
                {
                    var user = web.EnsureUser(admin.Name);
                    user.IsSiteAdmin = true;
                    user.Update();
                    web.Context.ExecuteQueryRetry();
                }

                if (siteSecurity.SiteSecurityPermissions != null)
                {
                    var existingRoleDefinitions = web.Context.LoadQuery(web.RoleDefinitions.Include(wr => wr.Name, wr => wr.BasePermissions, wr => wr.Description));
                    web.Context.ExecuteQueryRetry();

                    if (siteSecurity.SiteSecurityPermissions.RoleDefinitions.Any())
                    {
                        foreach (var templateRoleDefinition in siteSecurity.SiteSecurityPermissions.RoleDefinitions)
                        {
                            var siteRoleDefinition = existingRoleDefinitions.FirstOrDefault(erd => erd.Name == parser.ParseString(templateRoleDefinition.Name));
                            if (siteRoleDefinition == null)
                            {
                                scope.LogDebug("Creation role definition {0}", parser.ParseString(templateRoleDefinition.Name));
                                var roleDefinitionCI = new RoleDefinitionCreationInformation();
                                roleDefinitionCI.Name = parser.ParseString(templateRoleDefinition.Name);
                                roleDefinitionCI.Description = parser.ParseString(templateRoleDefinition.Description);
                                BasePermissions basePermissions = new BasePermissions();

                                foreach (var permission in templateRoleDefinition.Permissions)
                                {
                                    basePermissions.Set(permission);
                                }

                                roleDefinitionCI.BasePermissions = basePermissions;

                                web.RoleDefinitions.Add(roleDefinitionCI);
                                web.Context.ExecuteQueryRetry();
                            }
                            else
                            {
                                var isDirty = false;
                                if (siteRoleDefinition.Description != parser.ParseString(templateRoleDefinition.Description))
                                {
                                    siteRoleDefinition.Description = parser.ParseString(templateRoleDefinition.Description);
                                    isDirty = true;
                                }
                                var templateBasePermissions = new BasePermissions();
                                templateRoleDefinition.Permissions.ForEach(p => templateBasePermissions.Set(p));
                                if (siteRoleDefinition.BasePermissions != templateBasePermissions)
                                {
                                    isDirty = true;
                                    foreach (var permission in templateRoleDefinition.Permissions)
                                    {
                                        siteRoleDefinition.BasePermissions.Set(permission);
                                    }
                                }
                                if (isDirty)
                                {
                                    scope.LogDebug("Updating role definition {0}", parser.ParseString(templateRoleDefinition.Name));
                                    siteRoleDefinition.Update();
                                    web.Context.ExecuteQueryRetry();
                                }
                            }
                        }
                    }

                    var webRoleDefinitions = web.Context.LoadQuery(web.RoleDefinitions);
                    var groups = web.Context.LoadQuery(web.SiteGroups.Include(g => g.LoginName));
                    web.Context.ExecuteQueryRetry();

                    if (siteSecurity.SiteSecurityPermissions.RoleAssignments.Any())
                    {
                        foreach (var roleAssignment in siteSecurity.SiteSecurityPermissions.RoleAssignments)
                        {
                            Principal principal = groups.FirstOrDefault(g => g.LoginName == parser.ParseString(roleAssignment.Principal));
                            if (principal == null)
                            {
                                principal = web.EnsureUser(parser.ParseString(roleAssignment.Principal));
                            }

                            var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(web.Context);

                            var roleDefinition = webRoleDefinitions.FirstOrDefault(r => r.Name == roleAssignment.RoleDefinition);

                            if (roleDefinition != null)
                            {
                                roleDefinitionBindingCollection.Add(roleDefinition);
                            }
                            web.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
                            web.Context.ExecuteQueryRetry();
                        }
                    }
                }
            }
            return parser;
        }

        private static void AddUserToGroup(Web web, Group group, IEnumerable<User> members, PnPMonitoredScope scope)
        {
            if (members.Any())
            {
                scope.LogDebug("Adding users to group {0}", group.Title);
                try
                {
                    foreach (var user in members)
                    {
                        scope.LogDebug("Adding user {0}", user.Name);
                        var existingUser = web.EnsureUser(user.Name);
                        group.Users.AddUser(existingUser);

                    }
                    web.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_SiteSecurity_Add_users_failed_for_group___0_____1_____2_, group.Title, ex.Message, ex.StackTrace);
                    throw;
                }
            }
        }


        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // if this is a sub site then we're not creating security entities as by default security is inherited from the root site
                if (web.IsSubSite())
                {
                    return template;
                }

                web.Context.Load(web, w => w.HasUniqueRoleAssignments, w => w.Title);

                var ownerGroup = web.AssociatedOwnerGroup;
                var memberGroup = web.AssociatedMemberGroup;
                var visitorGroup = web.AssociatedVisitorGroup;
                web.Context.ExecuteQueryRetry();

                if (!ownerGroup.ServerObjectIsNull.Value)
                {
                    web.Context.Load(ownerGroup, o => o.Id, o => o.Users, o => o.Title);
                }
                if (!memberGroup.ServerObjectIsNull.Value)
                {
                    web.Context.Load(memberGroup, o => o.Id, o => o.Users, o => o.Title);
                }
                if (!visitorGroup.ServerObjectIsNull.Value)
                {
                    web.Context.Load(visitorGroup, o => o.Id, o => o.Users, o => o.Title);
                }
                web.Context.ExecuteQueryRetry();

                List<int> associatedGroupIds = new List<int>();
                var owners = new List<User>();
                var members = new List<User>();
                var visitors = new List<User>();
                if (!ownerGroup.ServerObjectIsNull.Value)
                {
                    associatedGroupIds.Add(ownerGroup.Id);
                    foreach (var member in ownerGroup.Users)
                    {
                        owners.Add(new User() { Name = member.LoginName });
                    }
                }
                if (!memberGroup.ServerObjectIsNull.Value)
                {
                    associatedGroupIds.Add(memberGroup.Id);
                    foreach (var member in memberGroup.Users)
                    {
                        members.Add(new User() { Name = member.LoginName });
                    }
                }
                if (!visitorGroup.ServerObjectIsNull.Value)
                {
                    associatedGroupIds.Add(visitorGroup.Id);
                    foreach (var member in visitorGroup.Users)
                    {
                        visitors.Add(new User() { Name = member.LoginName });
                    }
                }
                var siteSecurity = new SiteSecurity();
                siteSecurity.AdditionalOwners.AddRange(owners);
                siteSecurity.AdditionalMembers.AddRange(members);
                siteSecurity.AdditionalVisitors.AddRange(visitors);

                var query = from user in web.SiteUsers
                            where user.IsSiteAdmin
                            select user;
                var allUsers = web.Context.LoadQuery(query);

                web.Context.ExecuteQueryRetry();

                var admins = new List<User>();
                foreach (var member in allUsers)
                {
                    admins.Add(new User() { Name = member.LoginName });
                }
                siteSecurity.AdditionalAdministrators.AddRange(admins);

                if (creationInfo.IncludeSiteGroups)
                {
                    web.Context.Load(web.SiteGroups,
                        o => o.IncludeWithDefaultProperties(
                            gr => gr.Title,
                            gr => gr.AllowMembersEditMembership,
                            gr => gr.AutoAcceptRequestToJoinLeave,
                            gr => gr.AllowRequestToJoinLeave,
                            gr => gr.Description,
                            gr => gr.Users.Include(u => u.LoginName),
                            gr => gr.OnlyAllowMembersViewMembership,
                            gr => gr.Owner.LoginName,
                            gr => gr.RequestToJoinLeaveEmailSetting
                            ));

                    web.Context.ExecuteQueryRetry();

                    foreach (var group in web.SiteGroups.AsEnumerable().Where(o => !associatedGroupIds.Contains(o.Id)))
                    {
                        scope.LogDebug("Processing group {0}", group.Title);
                        var siteGroup = new SiteGroup()
                        {
                            Title = group.Title.Replace(web.Title, "{sitename}"),
                            AllowMembersEditMembership = group.AllowMembersEditMembership,
                            AutoAcceptRequestToJoinLeave = group.AutoAcceptRequestToJoinLeave,
                            AllowRequestToJoinLeave = group.AllowRequestToJoinLeave,
                            Description = group.Description,
                            OnlyAllowMembersViewMembership = group.OnlyAllowMembersViewMembership,
                            Owner = ReplaceGroupTokens(web, group.Owner.LoginName),
                            RequestToJoinLeaveEmailSetting = group.RequestToJoinLeaveEmailSetting
                        };

                        foreach (var member in group.Users)
                        {
                            siteGroup.Members.Add(new User() { Name = member.LoginName });
                        }
                        siteSecurity.SiteGroups.Add(siteGroup);
                    }
                }

                var webRoleDefinitions = web.Context.LoadQuery(web.RoleDefinitions.Include(r => r.Name, r => r.Description, r => r.BasePermissions, r => r.RoleTypeKind));
                web.Context.ExecuteQueryRetry();

                if (web.HasUniqueRoleAssignments)
                {
                    var permissionKeys = Enum.GetNames(typeof(PermissionKind));

                    foreach (var webRoleDefinition in webRoleDefinitions)
                    {
                        if (webRoleDefinition.RoleTypeKind == RoleType.None)
                        {
                            scope.LogDebug("Processing custom role definition {0}", webRoleDefinition.Name);
                            var modelRoleDefinitions = new Model.RoleDefinition();

                            modelRoleDefinitions.Description = webRoleDefinition.Description;
                            modelRoleDefinitions.Name = webRoleDefinition.Name;
                            var permissions = new List<PermissionKind>();

                            foreach (var permissionKey in permissionKeys)
                            {
                                var permissionKind = (PermissionKind)Enum.Parse(typeof(PermissionKind), permissionKey);
                                if (webRoleDefinition.BasePermissions.Has(permissionKind))
                                {
                                    modelRoleDefinitions.Permissions.Add(permissionKind);
                                }
                            }
                            siteSecurity.SiteSecurityPermissions.RoleDefinitions.Add(modelRoleDefinitions);
                        }
                        else
                        {
                            scope.LogDebug("Skipping OOTB role definition {0}", webRoleDefinition.Name);
                        }
                    }

                    var webRoleAssignments = web.Context.LoadQuery(web.RoleAssignments.Include(
                        r => r.RoleDefinitionBindings.Include(
                            rd => rd.Name,
                            rd => rd.RoleTypeKind),
                        r => r.Member.LoginName));

                    web.Context.ExecuteQueryRetry();

                    foreach (var webRoleAssignment in webRoleAssignments)
                    {
                        if (webRoleAssignment.Member.LoginName != "Excel Services Viewers")
                        {
                            foreach (var roleDefinition in webRoleAssignment.RoleDefinitionBindings)
                            {
                                if (roleDefinition.RoleTypeKind != RoleType.Guest)
                                {
                                    var modelRoleAssignment = new Model.RoleAssignment();
                                    modelRoleAssignment.RoleDefinition = roleDefinition.Name;
                                    modelRoleAssignment.Principal = ReplaceGroupTokens(web, webRoleAssignment.Member.LoginName);
                                    siteSecurity.SiteSecurityPermissions.RoleAssignments.Add(modelRoleAssignment);
                                }
                            }
                        }
                    }

                    template.Security = siteSecurity;
                }
                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);

                }
            }
            return template;
        }

        private string ReplaceGroupTokens(Web web, string loginName)
        {
            if (!web.AssociatedOwnerGroup.ServerObjectIsNull.Value)
            {
                loginName = loginName.Replace(web.AssociatedOwnerGroup.Title, "{associatedownergroup}");
            }
            if (!web.AssociatedMemberGroup.ServerObjectIsNull.Value)
            {
                loginName = loginName.Replace(web.AssociatedMemberGroup.Title, "{associatedmembergroup}");
            }
            if (!web.AssociatedVisitorGroup.ServerObjectIsNull.Value)
            {
                loginName = loginName.Replace(web.AssociatedVisitorGroup.Title, "{associatedvisitorgroup}");
            }
            return loginName;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            foreach (var user in baseTemplate.Security.AdditionalAdministrators)
            {
                int index = template.Security.AdditionalAdministrators.FindIndex(f => f.Name.Equals(user.Name));

                if (index > -1)
                {
                    template.Security.AdditionalAdministrators.RemoveAt(index);
                }
            }

            foreach (var user in baseTemplate.Security.AdditionalMembers)
            {
                int index = template.Security.AdditionalMembers.FindIndex(f => f.Name.Equals(user.Name));

                if (index > -1)
                {
                    template.Security.AdditionalMembers.RemoveAt(index);
                }
            }

            foreach (var user in baseTemplate.Security.AdditionalOwners)
            {
                int index = template.Security.AdditionalOwners.FindIndex(f => f.Name.Equals(user.Name));

                if (index > -1)
                {
                    template.Security.AdditionalOwners.RemoveAt(index);
                }
            }

            foreach (var user in baseTemplate.Security.AdditionalVisitors)
            {
                int index = template.Security.AdditionalVisitors.FindIndex(f => f.Name.Equals(user.Name));

                if (index > -1)
                {
                    template.Security.AdditionalVisitors.RemoveAt(index);
                }
            }

            foreach (var baseSiteGroup in baseTemplate.Security.SiteGroups)
            {
                var templateSiteGroup = template.Security.SiteGroups.FirstOrDefault(sg => sg.Title == baseSiteGroup.Title);
                if (templateSiteGroup != null)
                {
                    if (templateSiteGroup.Equals(baseSiteGroup))
                    {
                        template.Security.SiteGroups.Remove(templateSiteGroup);
                    }
                }
            }

            foreach (var baseRoleDef in baseTemplate.Security.SiteSecurityPermissions.RoleDefinitions)
            {
                var templateRoleDef = template.Security.SiteSecurityPermissions.RoleDefinitions.FirstOrDefault(rd => rd.Name == baseRoleDef.Name);
                if (templateRoleDef != null)
                {
                    if (templateRoleDef.Equals(baseRoleDef))
                    {
                        template.Security.SiteSecurityPermissions.RoleDefinitions.Remove(templateRoleDef);
                    }
                }
            }

            foreach (var baseRoleAssignment in baseTemplate.Security.SiteSecurityPermissions.RoleAssignments)
            {
                var templateRoleAssignments = template.Security.SiteSecurityPermissions.RoleAssignments.Where(ra => ra.Principal == baseRoleAssignment.Principal).ToList();
                foreach (var templateRoleAssignment in templateRoleAssignments)
                {
                    if (templateRoleAssignment.Equals(baseRoleAssignment))
                    {
                        template.Security.SiteSecurityPermissions.RoleAssignments.Remove(templateRoleAssignment);
                    }
                }
            }

            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = (template.Security.AdditionalAdministrators.Any() ||
                    template.Security.AdditionalMembers.Any() ||
                    template.Security.AdditionalOwners.Any() ||
                    template.Security.AdditionalVisitors.Any() ||
                    template.Security.SiteGroups.Any()) &&
                    !web.IsSubSite();
            }

            return _willProvision.Value;

        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = !web.IsSubSite();
            }
            return _willExtract.Value;
        }
    }
}
