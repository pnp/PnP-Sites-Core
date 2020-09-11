#if !ONPREMISES
#if NETSTANDARD2_0
using Microsoft.AspNetCore.StaticFiles;
#endif
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Object Handler to manage Microsoft Teams stuff
    /// </summary>
    internal class ObjectTeams : ObjectHierarchyHandlerBase
    {
        public override string Name => "Teams";

        /// <summary>
        /// Creates a new Team from a PnP Provisioning Schema definition
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="connector">The PnP File connector</param>
        /// <param name="team">The Team to provision</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The provisioned Team as a JSON object</returns>
        private static JToken CreateTeamFromProvisioningSchema(PnPMonitoredScope scope, TokenParser parser, FileConnectorBase connector, Team team, string accessToken)
        {
            string teamId = null;

            // If we have to Clone an existing Team
            if (!string.IsNullOrWhiteSpace(team.CloneFrom))
            {
                teamId = CloneTeam(scope, team, parser, accessToken);
            }
            // If we start from an already existing Group
            else if (!string.IsNullOrEmpty(team.GroupId))
            {
                // We need to parse the GroupId, if it is a token
                var parsedGroupId = parser.ParseString(team.GroupId);

                // Check if the Group exists
                if (GroupExistsById(scope, parsedGroupId, accessToken))
                {
                    // Then promote the Group into a Team or update it, if it already exists. Patching a team doesn't return an ID, so use the parsedGroupId directly (teamId and groupId are the same). 
                    teamId = CreateOrUpdateTeamFromGroup(scope, team, parser, parsedGroupId, accessToken) ?? parsedGroupId;
                }
                else
                {
                    // Log the exception and return NULL (i.e. cancel)
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_GroupDoesNotExists, parsedGroupId);
                    return null;
                }
            }
            // Otherwise create a Team from scratch
            else
            {
                teamId = CreateOrUpdateTeam(scope, team, parser, accessToken);
            }

            if (!string.IsNullOrEmpty(teamId))
            {
                // Wait to be sure that the Team is ready before configuring it
                WaitForTeamToBeReady(accessToken, teamId);

                // And now we configure security, channels, and apps
                // Only configure Security, if Security is configured
                if (team.Security != null)
                {
                    if (!SetGroupSecurity(scope, parser, team, teamId, accessToken)) return null;
                }
                if (!SetTeamChannels(scope, parser, team, teamId, accessToken)) return null;
                if (!SetTeamApps(scope, parser, team, teamId, accessToken)) return null;

                // So far the Team's photo cannot be set if we don't have an already existing mailbox
                if (!SetTeamPhoto(scope, parser, connector, team, teamId, accessToken)) return null;

                // Call Archive or Unarchive for the current Team
                ArchiveTeam(scope, teamId, team.Archived, accessToken);

                try
                {
                    // Get the whole Team that we just created and return it back as the method result
                    return JToken.Parse(HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}", accessToken));
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingError, ex.Message);
                }
            }

            return null;
        }

        private static void WaitForTeamToBeReady(string accessToken, string teamId)
        {
            // Wait for the Team to be ready
            bool wait = true;
            int iterations = 0;
            while (wait)
            {
                iterations++;

                try
                {
                    var jsonOwners = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/owners?$select=id", accessToken);
                    if (!string.IsNullOrEmpty(jsonOwners))
                    {
                        wait = false;
                    }
                }
                catch (Exception)
                {
                    // In case of exception wait for 5 secs
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(5));
                }

                // Don't wait more than 1 minute
                if (iterations > 12)
                {
                    wait = false;
                }
            }
        }

        private static string[] GetAllIdsForAllGroupsWithTeams(string accessToken)
        {
            var groupids = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=Id", accessToken);
            var value = JObject.Parse(groupids).Value<JArray>("value");
            return value.Select(t => t.Value<string>("id")).ToArray();
        }

        /// <summary>
        /// Checks if a Group exists by ID
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="groupId">The ID of the Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Group exists or not</returns>
        private static bool GroupExistsById(PnPMonitoredScope scope, string groupId, string accessToken)
        {
            var alreadyExistingGroupId = GraphHelper.ItemAlreadyExists($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups", "id", groupId, accessToken);
            return (alreadyExistingGroupId != null);
        }

        /// <summary>
        /// Checks if a Group exists by MailNickname
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="mailNickname">The ID of the Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The ID of an already existing Group with the provided MailNickname, if any</returns>
        private static string GetGroupIdByMailNickname(PnPMonitoredScope scope, string mailNickname, string accessToken)
        {
            var alreadyExistingGroupId = !string.IsNullOrEmpty(mailNickname) ?
                GraphHelper.ItemAlreadyExists($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups", "mailNickname", mailNickname, accessToken) :
                null;

            return (alreadyExistingGroupId);
        }

        /// <summary>
        /// Creates or updates a Team object via Graph
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="team">The Team to create</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The ID of the created or update Team</returns>
        private static string CreateOrUpdateTeam(PnPMonitoredScope scope, Team team, TokenParser parser, string accessToken)
        {
            var parsedMailNickname = !string.IsNullOrEmpty(team.MailNickname) ? parser.ParseString(team.MailNickname).ToLower() : null;

            if (string.IsNullOrEmpty(parsedMailNickname))
            {
                parsedMailNickname = CreateMailNicknameFromDisplayName(team.DisplayName);
            }

            // Check if the Group/Team already exists
            var alreadyExistingGroupId = GetGroupIdByMailNickname(scope, parsedMailNickname, accessToken);

            // If the Group already exists, we don't need to create it
            if (string.IsNullOrEmpty(alreadyExistingGroupId))
            {
                // Otherwise we create the Group, first

                // Prepare the IDs for owners and members
                string[] desiredOwnerIds;
                string[] desiredMemberIds;
                if (team.Security != null)
                {
                    try
                    {
                        var userIdsByUPN = team.Security.Owners
                            .Select(o => o.UserPrincipalName)
                            .Concat(team.Security.Members.Select(m => m.UserPrincipalName))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(k => k, k =>
                            {
                                var jsonUser = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{Uri.EscapeDataString(k.Replace("'", "''"))}?$select=id", accessToken);
                                return JToken.Parse(jsonUser).Value<string>("id");
                            });

                        desiredOwnerIds = team.Security.Owners.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
                        desiredMemberIds = team.Security.Members.Select(o => userIdsByUPN[o.UserPrincipalName]).Union(desiredOwnerIds).ToArray();
                    }
                    catch (Exception ex)
                    {
                        scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingUserError, ex.Message);
                        return (null);
                    }
                }
                else
                {
                    desiredOwnerIds = new string[0];
                    desiredMemberIds = new string[0];
                }

                var groupCreationRequest = new
                {
                    displayName = parser.ParseString(team.DisplayName),
                    description = parser.ParseString(team.Description),
                    groupTypes = new string[]
                    {
                        "Unified"
                    },
                    mailEnabled = true,
                    mailNickname = parsedMailNickname,
                    securityEnabled = false,
                    visibility = team.Visibility.ToString(),
                    owners_odata_bind = (from o in desiredOwnerIds select $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{Uri.EscapeDataString(o.Replace("'", "''"))}").ToArray(),
                    members_odata_bind = (from m in desiredMemberIds select $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{Uri.EscapeDataString(m.Replace("'", "''"))}").ToArray()
                };

                // Make the Graph request to create the Office 365 Group
                var createdGroupJson = HttpHelper.MakePostRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups",
                    groupCreationRequest, HttpHelper.JsonContentType, accessToken);
                var createdGroupId = JToken.Parse(createdGroupJson).Value<string>("id");

                // Wait for the Group to be ready
                bool wait = true;
                int iterations = 0;
                while (wait)
                {
                    iterations++;

                    try
                    {
                        var jsonGroup = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{createdGroupId}", accessToken);
                        if (!string.IsNullOrEmpty(jsonGroup))
                        {
                            wait = false;
                        }
                    }
                    catch (Exception)
                    {
                        // In case of exception wait for 5 secs
                        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(5));
                    }

                    // Don't wait more than 1 minute
                    if (iterations > 12)
                    {
                        wait = false;
                    }
                }

                team.GroupId = createdGroupId;
            }
            else
            {
                // Otherwise use the already existing Group ID
                team.GroupId = alreadyExistingGroupId;
            }

            // Then we Teamify the Group
            return CreateOrUpdateTeamFromGroup(scope, team, parser, team.GroupId, accessToken);
        }

        /// <summary>
        /// Creates a Team object via Graph cloning an already existing one
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="team">The Team to create</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The ID of the created Team</returns>
        private static string CloneTeam(PnPMonitoredScope scope, Team team, TokenParser parser, string accessToken)
        {
            var content = PrepareTeamCloneRequestContent(team, parser);

            var teamId = GraphHelper.CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST_WITH_RESPONSE_HEADERS,
                $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{parser.ParseString(team.CloneFrom)}/clone",
                content,
                HttpHelper.JsonContentType,
                accessToken,
                "Conflict",
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_AlreadyExists,
                "id",
                team.GroupId,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError,
                canPatch: true);

            return (teamId);
        }

        /// <summary>
        /// Prepares the JSON object for the request to clone a Team
        /// </summary>
        /// <param name="team">The Domain Model Team object</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <returns>The JSON object ready to be serialized into the JSON request</returns>
        private static object PrepareTeamCloneRequestContent(Team team, TokenParser parser)
        {
            var content = new
            {
                DisplayName = parser.ParseString(team.DisplayName),
                Description = parser.ParseString(team.Description),
                Classification = parser.ParseString(team.Classification),
                Mailnickname = parser.ParseString(team.MailNickname),
                team.Visibility,
                partsToClone = "apps,tabs,settings,channels,members", // Clone everything
            };

            return (content);
        }

        /// <summary>
        /// Creates or updates a Team object via Graph promoting an existing Group
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="team">The Team to create</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="groupId">The ID of the Group to promote into a Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The ID of the created or updated Team</returns>
        private static string CreateOrUpdateTeamFromGroup(PnPMonitoredScope scope, Team team, TokenParser parser, string groupId, string accessToken)
        {
            bool isCurrentlyArchived = false;
            try
            {
                // Check the archival status of the team
                string archiveStatusReq = HttpHelper.MakeGetRequestForString(
                    $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{groupId}?$select=isArchived", accessToken: accessToken);

                isCurrentlyArchived = JToken.Parse(archiveStatusReq).Value<bool>("isArchived");
            }
            catch(Exception ex)
            {
                scope.LogError("Error checking archive status", ex.Message);
            }            

            // If the Team is currently archived
            if (isCurrentlyArchived)
            {
                // and if the templates declares to have it unarchived
                if (!team.Archived)
                {
                    // first unarchive Team because we have set the flag to false
                    HttpHelper.MakePostRequest(
                        $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{groupId}/unarchive", accessToken: accessToken);
                }
                else
                {
                    // Else, we will skip processing the team
                    scope.LogWarning($"Team {team.DisplayName} is currently archived, so processing it will be skipped");
                    return null;
                }
            }

            // Now process the Team create or update request
            return CreateOrUpdateTeamFromGroupInternal(scope, team, parser, groupId, accessToken);
        }

        private static string CreateOrUpdateTeamFromGroupInternal(PnPMonitoredScope scope, Team team, TokenParser parser, string groupId, string accessToken)
        {
            var content = PrepareTeamRequestContent(team, parser);

            bool wait = true;
            int iterations = 0;
            string teamId = null;
            while (wait)
            {
                iterations++;

                try
                {
                    teamId = GraphHelper.CreateOrUpdateGraphObject(scope,
                    HttpMethodVerb.PUT,
                    $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{groupId}/team",
                    content,
                    HttpHelper.JsonContentType,
                    accessToken,
                    "Conflict",
                    CoreResources.Provisioning_ObjectHandlers_Teams_Team_AlreadyExists,
                    "id",
                    parser.ParseString(team.GroupId),
                    CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError,
                    canPatch: true);

                    wait = false;
                }
                catch (Exception)
                {
                    // In case of exception wait for 5 secs
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(5));
                }

                // Don't wait more than 60 seconds
                if (iterations > 12)
                {
                    wait = false;
                }
            }
            return (teamId);
        }

        /// <summary>
        /// Prepares the JSON object for the request to create/update a Team
        /// </summary>
        /// <param name="team">The Domain Model Team object</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <returns>The JSON object ready to be serialized into the JSON request</returns>
        private static object PrepareTeamRequestContent(Team team, TokenParser parser)
        {
            var content = new
            {
                //template_odata_bind = $"{GraphHelper.MicrosoftGraphBaseURI}beta/teamsTemplates('standard')",
                //DisplayName = parser.ParseString(team.DisplayName),
                //Description = parser.ParseString(team.Description),
                //Classification = parser.ParseString(team.Classification),
                //Mailnickname = parser.ParseString(team.MailNickname),
                //team.Specialization,
                //team.Visibility,
                funSettings = new
                {
                    team.FunSettings?.AllowGiphy,
                    team.FunSettings?.GiphyContentRating,
                    team.FunSettings?.AllowStickersAndMemes,
                    team.FunSettings?.AllowCustomMemes,
                },
                guestSettings = new
                {
                    team.GuestSettings?.AllowCreateUpdateChannels,
                    team.GuestSettings?.AllowDeleteChannels,
                },
                memberSettings = new
                {
                    team.MemberSettings?.AllowCreateUpdateChannels,
                    team.MemberSettings?.AllowAddRemoveApps,
                    team.MemberSettings?.AllowDeleteChannels,
                    team.MemberSettings?.AllowCreateUpdateRemoveTabs,
                    team.MemberSettings?.AllowCreateUpdateRemoveConnectors,
                    team.MemberSettings?.AllowCreatePrivateChannels,
                },
                messagingSettings = new
                {
                    team.MessagingSettings?.AllowUserEditMessages,
                    team.MessagingSettings?.AllowUserDeleteMessages,
                    team.MessagingSettings?.AllowOwnerDeleteMessages,
                    team.MessagingSettings?.AllowTeamMentions,
                    team.MessagingSettings?.AllowChannelMentions
                }
            };

            return (content);
        }

        /// <summary>
        /// Creates or updates a Team object via Graph
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="archived">A flag to declare to archive or unarchive the Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        private static void ArchiveTeam(PnPMonitoredScope scope, string teamId, bool archived, string accessToken)
        {
            string archiveStatusRequest = HttpHelper.MakeGetRequestForString(
                $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}?$select=isArchived", accessToken: accessToken);

            bool isCurrentlyArchived = JToken.Parse(archiveStatusRequest).Value<bool>("isArchived");

            try
            {
                if (archived && !isCurrentlyArchived)
                {
                    // Archive the Team
                    HttpHelper.MakePostRequest(
                        $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/archive", accessToken: accessToken);
                }
                else if (!archived && isCurrentlyArchived)
                {
                    // Unarchive the Team
                    HttpHelper.MakePostRequest(
                        $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/unarchive", accessToken: accessToken);
                }
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FailedArchiveUnarchive, teamId, ex.Message);
            }
        }

        /// <summary>
        /// Synchronizes Owners and Members with Team settings
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="team">The Team settings, including security settings</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Security settings have been provisioned or not</returns>
        private static bool SetGroupSecurity(PnPMonitoredScope scope, TokenParser parser, Team team, string teamId, string accessToken)
        {
            SetAllowToAddGuestsSetting(scope, teamId, team.Security.AllowToAddGuests, accessToken);

            string[] desideredOwnerIds;
            string[] desideredMemberIds;
            string[] finalOwnerIds;
            try
            {
                var userIdsByUPN = team.Security.Owners
                    .Select(o => o.UserPrincipalName)
                    .Concat(team.Security.Members.Select(m => m.UserPrincipalName))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(k => k, k =>
                    {
                        var parsedUser = parser.ParseString(k);
                        var jsonUser = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{Uri.EscapeDataString(parsedUser.Replace("'", "''"))}?$select=id", accessToken);
                        return JToken.Parse(jsonUser).Value<string>("id");
                    });

                desideredOwnerIds = team.Security.Owners.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
                desideredMemberIds = team.Security.Members.Select(o => userIdsByUPN[o.UserPrincipalName]).Union(desideredOwnerIds).ToArray();
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingUserError, ex.Message);
                return false;
            }

            string[] ownerIdsToAdd;
            string[] ownerIdsToRemove;
            try
            {
                // Get current group owners
                var jsonOwners = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/owners?$select=id", accessToken);

                string[] currentOwnerIds = GraphHelper.GetIdsFromList(jsonOwners);

                // Exclude owners already into the group
                ownerIdsToAdd = desideredOwnerIds.Except(currentOwnerIds).ToArray();

                if (team.Security.ClearExistingOwners)
                {
                    ownerIdsToRemove = currentOwnerIds.Except(desideredOwnerIds).ToArray();
                }
                else
                {
                    ownerIdsToRemove = new string[0];
                }

                // Define the complete set of owners
                finalOwnerIds = currentOwnerIds.Union(ownerIdsToAdd).Except(ownerIdsToRemove).ToArray();
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_ListingOwnersError, ex.Message);
                return false;
            }

            // Add new owners
            foreach (string ownerId in ownerIdsToAdd)
            {
                try
                {
                    object content = new JObject
                    {
                        ["@odata.id"] = $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{ownerId}"
                    };
                    HttpHelper.MakePostRequest($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/owners/$ref", content, "application/json", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_AddingOwnerError, ex.Message);
                    return false;
                }
            }

            // Remove exceeding owners
            foreach (string ownerId in ownerIdsToRemove)
            {
                try
                {
                    HttpHelper.MakeDeleteRequest($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/owners/{ownerId}/$ref", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingOwnerError, ex.Message);
                    return false;
                }
            }

            string[] memberIdsToAdd;
            string[] memberIdsToRemove;
            try
            {
                // Get current group members
                var jsonMembers = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/members?$select=id", accessToken);

                string[] currentMemberIds = GraphHelper.GetIdsFromList(jsonMembers);

                // Exclude members already into the group
                memberIdsToAdd = desideredMemberIds.Except(currentMemberIds).ToArray();

                if (team.Security.ClearExistingMembers)
                {
                    memberIdsToRemove = currentMemberIds.Except(desideredMemberIds.Union(finalOwnerIds)).ToArray();
                }
                else
                {
                    memberIdsToRemove = new string[0];
                }
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_ListingMembersError, ex.Message);
                return false;
            }

            // Add new members
            foreach (string memberId in memberIdsToAdd)
            {
                try
                {
                    object content = new JObject
                    {
                        ["@odata.id"] = $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{memberId}"
                    };
                    HttpHelper.MakePostRequest($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/members/$ref", content, "application/json", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_AddingMemberError, ex.Message);
                    return false;
                }
            }

            // Remove exceeding members
            foreach (string memberId in memberIdsToRemove)
            {
                try
                {
                    HttpHelper.MakeDeleteRequest($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/members/{memberId}/$ref", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingMemberError, ex.Message);
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Checks if the AllowToAddGuest setting already exists for the team connected unified group, and based on the outcome either creates or updates the setting.
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="allowToAddGuests">Boolean value indicating whether external sharing should be allowed or not.</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        private static void SetAllowToAddGuestsSetting(PnPMonitoredScope scope, string teamId, bool allowToAddGuests, string accessToken)
        {
            if (GetAllowToAddGuestsSetting(scope, teamId, accessToken))
            {
                UpdateAllowToAddGuestsSetting(scope, teamId, allowToAddGuests, accessToken);
            }
            else
            {
                CreateAllowToAddGuestsSetting(scope, teamId, allowToAddGuests, accessToken);
            }
        }

        /// <summary>
        /// Gets the AllowToAddGuests setting JSON (name and value) of the team connected unified group.
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>JSON object with name and value properties</returns>
        internal static bool GetAllowToAddGuestsSetting(PnPMonitoredScope scope, string teamId, string accessToken)
        {
            try
            {
                var groupGuestSettings = GetGroupUnifiedGuestSettings(scope, teamId, accessToken);
                if (groupGuestSettings != null && groupGuestSettings["values"] != null && groupGuestSettings["values"].FirstOrDefault(x => x["name"].Value<string>().Equals("AllowToAddGuests")) != null)
                {
                    return groupGuestSettings["values"].First(x => x["name"].ToString() == "AllowToAddGuests").Value<bool>();
                }
                return false;
            }
            catch (Exception e)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingMemberError, e.Message);
                return false;
            }
        }

        /// <summary>
        /// Gets the Group.Unified.Guest settings for the unified group that is connected to the team.
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>All guest related settings for the team connected unified group (not just external sharing)</returns>
        private static JToken GetGroupUnifiedGuestSettings(PnPMonitoredScope scope, string teamId, string accessToken)
        {
            try
            {
                var response = JToken.Parse(HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/settings", accessToken));
                return response["value"]?.FirstOrDefault(x => x["templateId"].ToString() == "08d542b9-071f-4e16-94b0-74abb372e3d9");
            }
            catch (Exception e)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingMemberError, e.Message);
                return null;
            }
        }

        /// <summary>
        /// Creates the AllowToAddGuests setting for the team connected unified group, and sets its value.
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="allowToAddGuests">Boolean value indicating whether external sharing should be allowed or not.</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        private static void CreateAllowToAddGuestsSetting(PnPMonitoredScope scope, string teamId, bool allowToAddGuests, string accessToken)
        {
            try
            {
                var body = $"{{'displayName': 'Group.Unified.Guest', 'templateId': '08d542b9-071f-4e16-94b0-74abb372e3d9', 'values': [{{'name': 'AllowToAddGuests','value': '{allowToAddGuests}'}}] }}";
                HttpHelper.MakePostRequest($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/settings", body, "application/json", accessToken);
            }
            catch (Exception e)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingMemberError, e.Message);
            }
        }

        /// <summary>
        /// Updates an existing AllowToAddGuests setting for the team connected unified group.
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="allowToAddGuests">Boolean value indicating whether external sharing should be allowed or not.</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        private static void UpdateAllowToAddGuestsSetting(PnPMonitoredScope scope, string teamId, bool allowToAddGuests, string accessToken)
        {
            try
            {
                var groupGuestSettings = GetGroupUnifiedGuestSettings(scope, teamId, accessToken);
                groupGuestSettings["values"].FirstOrDefault(x => x["name"].ToString() == "AllowToAddGuests")["value"] = allowToAddGuests.ToString();

                HttpHelper.MakePatchRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/settings/{groupGuestSettings["id"]}", groupGuestSettings, "application/json", accessToken);
            }
            catch (Exception e)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingMemberError, e.Message);
            }
        }

        /// <summary>
        /// Synchronizes Team Channels settings
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="team">The Team settings, including security settings</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Channels have been provisioned or not</returns>
        private static bool SetTeamChannels(PnPMonitoredScope scope, TokenParser parser, Team team, string teamId, string accessToken)
        {
            if (team.Channels == null) return true;

            var existingChannels = GetExistingTeamChannels(teamId, accessToken);

            foreach (var channel in team.Channels)
            {
                var existingChannel = existingChannels.FirstOrDefault(x => x["displayName"].ToString() == channel.DisplayName);

                var channelId = existingChannel == null ? CreateTeamChannel(scope, channel, teamId, accessToken) : UpdateTeamChannel(channel, teamId, existingChannel, accessToken);

                if (channelId == null) return false;

                if (channel.Tabs != null && channel.Tabs.Any())
                {
                    if (!SetTeamTabs(scope, parser, channel.Tabs, teamId, channelId, accessToken)) return false;
                }

                // TODO: Handle TabResources
                // We need to define a "schema" for their settings

                if (channel.Messages != null && channel.Messages.Any())
                {
                    if (!SetTeamChannelMessages(scope, parser, channel.Messages, teamId, channelId, accessToken)) return false;
                }
            }

            return true;
        }

        public static JToken GetExistingTeamChannels(string teamId, string accessToken)
        {
            return JToken.Parse(HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/channels", accessToken))["value"];
        }

        private static string UpdateTeamChannel(Model.Teams.TeamChannel channel, string teamId, JToken existingChannel, string accessToken)
        {
            // Not supported to update 'General' Channel
            if (channel.DisplayName.Equals("General", StringComparison.InvariantCultureIgnoreCase))
                return existingChannel["id"].ToString();

            var channelId = existingChannel["id"].ToString();
            var channelDisplayName = existingChannel["displayName"].ToString();
            var identicalChannelName = channel.DisplayName == channelDisplayName;

            // Prepare the request body for the Channel update
            var channelToUpdate = new
            {
                description = channel.Description,
                // You can't update a channel if its displayName is exactly the same, so remove it temporarily.
                displayName = identicalChannelName ? null : channel.DisplayName,
            };

            // Updating isFavouriteByDefault is currently not supported on either endpoint. Using the beta endpoint results in an error.
            HttpHelper.MakePatchRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/channels/{channelId}", channelToUpdate, HttpHelper.JsonContentType, accessToken);

            return channelId;
        }

        private static string CreateTeamChannel(PnPMonitoredScope scope, Model.Teams.TeamChannel channel, string teamId, string accessToken)
        {
            // Temporary variable, just in case
            List<String> channelMembers = null;

            if (channel.Private)
            {
                // Get the team owners, who will be set as members of the private channel
                // if the channel is private
                var teamOwnersString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/groups/{teamId}/owners", accessToken);
                channelMembers = new List<String>();

                foreach (var user in JObject.Parse(teamOwnersString)["value"] as JArray)
                {
                    channelMembers.Add((string)user["id"]);
                }
            }

            var channelToCreate = new
            {
                channel.Description,
                channel.DisplayName,
                channel.IsFavoriteByDefault,
                membershipType = channel.Private ? "private" : "standard",
                members = (channel.Private && channelMembers != null) ? (from m in channelMembers
                                                                         select new
                                                                         {
                                                                             private_channel_member_odata_type = "#microsoft.graph.aadUserConversationMember",
                                                                             private_channel_user_odata_bind = $"https://graph.microsoft.com/beta/users('{m}')",
                                                                             roles = new String[] { "owner" }
                                                                         }).ToArray() : null
            };

            var channelId = GraphHelper.CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST,
                $"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels",
                channelToCreate,
                HttpHelper.JsonContentType,
                accessToken,
                "NameAlreadyExists",
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_ChannelAlreadyExists,
                "displayName",
                channel.DisplayName,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError,
                false);

            return channelId;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="tabs">A collection of Tabs to be created or updated</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="channelId">the ID of the target Channel</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns></returns>
        public static bool SetTeamTabs(PnPMonitoredScope scope, TokenParser parser, TeamTabCollection tabs, string teamId, string channelId, string accessToken)
        {
            var existingTabs = GetExistingTeamChannelTabs(teamId, channelId, accessToken);

            foreach (var tab in tabs)
            {
                // Avoid ActivityLimitReached 
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(5));

                var existingTab = existingTabs.FirstOrDefault(x => x["displayName"] != null && HttpUtility.UrlDecode(x["displayName"].ToString()) == tab.DisplayName && x["teamsAppId"] != null && x["teamsAppId"].ToString() == tab.TeamsAppId);

                var tabId = existingTab == null ? CreateTeamTab(scope, tab, parser, teamId, channelId, accessToken) : UpdateTeamTab(tab, parser, teamId, channelId, existingTab["id"].ToString(), accessToken);

                if (tabId == null) return false;
            }
            if (tabs.Any())
            {
                // is there a wiki tab and not a newly created tab?
                var wikiTab = existingTabs.FirstOrDefault(x => x["teamsAppId"].Value<string>() == "com.microsoft.teamspace.tab.wiki");
                if (wikiTab != null && tabs.FirstOrDefault(t => t.TeamsAppId == "com.microsoft.teamspace.tab.wiki") == null)
                {
                    RemoveTeamTab(wikiTab["id"].Value<string>(), channelId, teamId, accessToken);
                }
            }
            return true;
        }

        private static void RemoveTeamTab(string tabId, string channelId, string teamId, string accessToken)
        {
            HttpHelper.MakeDeleteRequest($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/channels/{channelId}/tabs/{tabId}", accessToken);
        }

        public static JToken GetExistingTeamChannelTabs(string teamId, string channelId, string accessToken)
        {
            return JToken.Parse(HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/tabs", accessToken))["value"];
        }

        private static string UpdateTeamTab(TeamTab tab, TokenParser parser, string teamId, string channelId, string tabId, string accessToken)
        {
            var displayname = parser.ParseString(tab.DisplayName);

            if (!tab.Remove)
            {
                // teamsAppId is not allowed in the request
                var teamsAppId = parser.ParseString(tab.TeamsAppId);
                tab.TeamsAppId = null;

                if (tab.Configuration != null)
                {
                    tab.Configuration.EntityId = parser.ParseString(tab.Configuration.EntityId);
                    tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                    tab.Configuration.RemoveUrl = parser.ParseString(tab.Configuration.RemoveUrl);
                    tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.WebsiteUrl);
                }


                // Prepare the request body for the Tab update
                var tabToUpdate = new
                {
                    displayName = displayname,
                    configuration = tab.Configuration != null
                        ? new
                        {
                            tab.Configuration.EntityId,
                            tab.Configuration.ContentUrl,
                            tab.Configuration.RemoveUrl,
                            tab.Configuration.WebsiteUrl
                        } : null,
                };

            HttpHelper.MakePatchRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/channels/{channelId}/tabs/{tabId}", tabToUpdate, HttpHelper.JsonContentType, accessToken);

                // Add the teamsAppId back now that we've updated the tab
                tab.TeamsAppId = teamsAppId;
            }
            else
            {
                // Simply delete the tab
                HttpHelper.MakeDeleteRequest($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/tabs/{tabId}", accessToken);
            }

            return tabId;
        }

        private static string CreateTeamTab(PnPMonitoredScope scope, TeamTab tab, TokenParser parser, string teamId, string channelId, string accessToken)
        {
            // There is no reason to create a tab that has to be removed
            if (tab.Remove)
            {
                return null;
            }

            var displayname = parser.ParseString(tab.DisplayName);
            var teamsAppId = parser.ParseString(tab.TeamsAppId);

            if (tab.Configuration != null)
            {
                // https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs
                switch (tab.TeamsAppId)
                {
                    case "com.microsoft.teamspace.tab.web": // Website
                        {
                            tab.Configuration.EntityId = null;
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = null;
                            tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.WebsiteUrl);
                            break;
                        }
                    case "com.microsoft.teamspace.tab.planner": // Planner
                        {
                            tab.Configuration.EntityId = parser.ParseString(tab.Configuration.EntityId);
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            break;
                        }
                    case "com.microsoftstream.embed.skypeteamstab": // Stream
                        {
                            tab.Configuration.EntityId = null;
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = null;
                            tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.WebsiteUrl);
                            break;
                        }
                    case "81fef3a6-72aa-4648-a763-de824aeafb7d": // Forms
                        {
                            tab.Configuration.EntityId = parser.ParseString(tab.Configuration.EntityId);
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = null;
                            tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.WebsiteUrl);
                            break;
                        }
                    case "com.microsoft.teamspace.tab.file.staticviewer.word": // Word
                    case "com.microsoft.teamspace.tab.file.staticviewer.excel": // Excel
                    case "com.microsoft.teamspace.tab.file.staticviewer.powerpoint": // PowerPoint
                    case "com.microsoft.teamspace.tab.file.staticviewer.pdf": // PDF
                        {
                            tab.Configuration.EntityId = parser.ParseString(tab.Configuration.EntityId);
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = null;
                            tab.Configuration.WebsiteUrl = null;
                            break;
                        }
                    case "com.microsoft.teamspace.tab.wiki": // Wiki, no configuration possible
                        {
                            tab.Configuration = null;
                            break;
                        }
                    case "com.microsoft.teamspace.tab.files.sharepoint": // Document library
                        {
                            tab.Configuration.EntityId = "";
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = null;
                            tab.Configuration.WebsiteUrl = null;
                            break;
                        }
                    case "0d820ecd-def2-4297-adad-78056cde7c78": // OneNote
                        {
                            tab.Configuration.EntityId = parser.ParseString(tab.Configuration.EntityId);
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = parser.ParseString(tab.Configuration.RemoveUrl);
                            tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.WebsiteUrl);
                            break;
                        }
                    case "com.microsoft.teamspace.tab.powerbi": //  Power BI
                        {
                            tab.Configuration = null;
                            break;
                        }
                    case "2a527703-1f6f-4559-a332-d8a7d288cd88": // SharePoint page and lists
                        {
                            tab.Configuration = null;
                            break;
                        }
                    default:
                        {
                            tab.Configuration.EntityId = parser.ParseString(tab.Configuration.EntityId);
                            tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                            tab.Configuration.RemoveUrl = parser.ParseString(tab.Configuration.RemoveUrl);
                            tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.WebsiteUrl);
                            break;
                        }
                }

            }

            Dictionary<string, object> tabToCreate = new Dictionary<string, object>();
            tabToCreate.Add("displayName", displayname);
            tabToCreate.Add("configuration", tab.Configuration != null
                                        ? new
                                        {
                                            tab.Configuration.EntityId,
                                            tab.Configuration.ContentUrl,
                                            tab.Configuration.RemoveUrl,
                                            tab.Configuration.WebsiteUrl
                                        }
                                        : null);
            tabToCreate.Add("teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + teamsAppId);

            var tabId = GraphHelper.CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST,
                $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/channels/{channelId}/tabs",
                JsonConvert.SerializeObject(tabToCreate),
                HttpHelper.JsonContentType,
                accessToken,
                "NameAlreadyExists",
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_TabAlreadyExists,
                "displayName",
                displayname,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError,
                false);

            return tabId;
        }

        public static bool SetTeamChannelMessages(PnPMonitoredScope scope, TokenParser parser, TeamChannelMessageCollection messages, string teamId, string channelId, string accessToken)
        {
            foreach (var message in messages)
            {
                var messageId = CreateTeamChannelMessage(scope, parser, message, teamId, channelId, accessToken);
                if (messageId == null) return false;
            }

            return true;
        }

        private static JObject CleanUpMessage(JObject message)
        {
            List<string> propertiesToRemove = new List<string> { "createdDateTime", "id", "webUrl" };
            foreach (var property in propertiesToRemove)
            {
                message.Remove(property);
            }
            return message;
        }

        private static string CreateTeamChannelMessage(PnPMonitoredScope scope, TokenParser parser, TeamChannelMessage message, string teamId, string channelId, string accessToken)
        {
            var messageString = parser.ParseString(message.Message);
            var messageObject = default(JObject);

            try
            {
                // If the message is already in JSON format, we just use it
                messageObject = JObject.Parse(messageString);
            }
            catch
            {
                // Otherwise try to build the JSON message content from scratch
                messageObject = JObject.Parse($"{{ \"body\": {{ \"content\": \"{messageString}\" }} }}");
            }

            // We cannot set the createdDateTime value when posting a message.
            messageObject = CleanUpMessage(messageObject);

            var messageId = GraphHelper.CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST,
                $"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/messages",
                messageObject,
                HttpHelper.JsonContentType,
                accessToken,
                null,
                null,
                null,
                null,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_CannotSendMessage,
                false);

            return messageId;
        }

        /// <summary>
        /// Synchronizes Team Apps settings
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">Token parser</param>
        /// <param name="team">The Team settings, including security settings</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Apps have been provisioned or not</returns>
        private static bool SetTeamApps(PnPMonitoredScope scope, TokenParser parser, Team team, string teamId, string accessToken)
        {
            foreach (var app in team.Apps)
            {
                object appToCreate = new JObject
                {
                    ["teamsApp@odata.bind"] = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + parser.ParseString(app.AppId)
                };

                var id = GraphHelper.CreateOrUpdateGraphObject(scope,
                    HttpMethodVerb.POST,
                    $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{teamId}/installedApps",
                    appToCreate,
                    HttpHelper.JsonContentType,
                    accessToken,
                    null,
                    null,
                    null,
                    null,
                    CoreResources.Provisioning_ObjectHandlers_Teams_Team_AppProvisioningError,
                    canPatch: false);
            }

            return true;
        }

        /// <summary>
        /// Synchronizes Team's Photo
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="connector">The PnP File Connector</param>
        /// <param name="team">The Team settings, including security settings</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Apps have been provisioned or not</returns>
        private static bool SetTeamPhoto(PnPMonitoredScope scope, TokenParser parser, FileConnectorBase connector, Team team, string teamId, string accessToken)
        {
            if (!string.IsNullOrEmpty(team.Photo) && connector != null)
            {
                var photoPath = parser.ParseString(team.Photo);
                var photoBytes = ConnectorFileHelper.GetFileBytes(connector, team.Photo);

                using (var photoStream = new MemoryStream(photoBytes))
                {
#if !NETSTANDARD2_0
                    var contentType = MimeMapping.GetMimeMapping(photoPath);
#else
                    string contentType;
                    new FileExtensionContentTypeProvider().TryGetContentType(photoPath, out contentType);
                    if (contentType == null)
                    {
                        contentType = "application/octet-stream";
                    }
#endif
                    int maxRetries = 10;
                    int retry = 0;
                    while (retry < maxRetries)
                        try
                        {
                            HttpHelper.MakePutRequest(
                                $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/photo/$value",
                                photoStream, contentType, accessToken);
                            break;
                        }
                        catch (Exception)
                        {
                            retry++;
                            Thread.Sleep(5000 * retry); // wait
                        }
                }
            }
            return true;
        }

        /// <summary>
        /// Creates a Team starting from a JSON template
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="teamTemplate">The Team JSON template</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The provisioned Team as a JSON object</returns>
        private static JToken CreateTeamFromJsonTemplate(PnPMonitoredScope scope, TokenParser parser, TeamTemplate teamTemplate, string accessToken)
        {
            HttpResponseHeaders responseHeaders;
            try
            {
                var content = OverwriteJsonTemplateProperties(parser, teamTemplate);
                responseHeaders = HttpHelper.MakePostRequestForHeaders($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams", content, "application/json", accessToken);
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_ProvisioningError, ex.Message);
                return null;
            }

            bool wait = true;
            int iterations = 0;
            while (wait)
            {
                iterations++;

                try
                {
                    var teamId = responseHeaders.Location.ToString().Split('\'')[1];
                    var team = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}", accessToken);
                    wait = false;
                    return JToken.Parse(team);
                }
                catch (Exception)
                {
                    // In case of exception wait for 10 secs
                    Thread.Sleep(TimeSpan.FromSeconds(10));
                }

                // Don't wait more than 1 minute
                if (iterations > 6)
                {
                    wait = false;
                }                
            }

            return null;
        }

        /// <summary>
        /// Allows to overwrite some settings of the templates provisioned through JSON template
        /// </summary>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="teamTemplate">The Team JSON template</param>
        /// <returns>The updated Team JSON template</returns>
        private static string OverwriteJsonTemplateProperties(TokenParser parser, TeamTemplate teamTemplate)
        {
            var jsonTemplate = parser.ParseString(teamTemplate.JsonTemplate);
            var team = JToken.Parse(jsonTemplate);

            if (teamTemplate.DisplayName != null) team["displayName"] = teamTemplate.DisplayName;
            if (teamTemplate.Description != null) team["description"] = teamTemplate.Description;
            if (!string.IsNullOrEmpty(teamTemplate.Classification)) team["classification"] = teamTemplate.Classification;
            team["visibility"] = teamTemplate.Visibility.ToString();

            return team.ToString();
        }

        #region PnP Provisioning Engine infrastructural code

        public override bool WillProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = hierarchy.Teams?.TeamTemplates?.Any() |
                    hierarchy.Teams?.Teams?.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ExtractConfiguration configuration)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }

        public override TokenParser ProvisionObjects(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser, ApplyConfiguration configuration)
        {
            using (var scope = new PnPMonitoredScope(Name))
            {
                int totalCount = (hierarchy.Teams.TeamTemplates != null ? hierarchy.Teams.TeamTemplates.Count : 0) + (hierarchy.Teams.Teams != null ? hierarchy.Teams.Teams.Count : 0);
                int currentProgress = 0;
                // Prepare a method global variable to store the Access Token
                string accessToken = null;

                // - Teams based on JSON templates
                var teamTemplates = hierarchy.Teams?.TeamTemplates;
                if (teamTemplates != null && teamTemplates.Any())
                {
                    foreach (var teamTemplate in teamTemplates)
                    {
                        WriteSubProgress("Teams", "Team Template", currentProgress, totalCount);
                        if (PnPProvisioningContext.Current != null)
                        {
                            // Get a fresh Access Token for every request
                            accessToken = PnPProvisioningContext.Current.AcquireToken(new Uri(GraphHelper.MicrosoftGraphBaseURI).Authority, "Group.ReadWrite.All");

                            if (accessToken != null)
                            {
                                // Create the Team starting from the JSON template
                                var team = CreateTeamFromJsonTemplate(scope, parser, teamTemplate, accessToken);

                                // TODO: possible further processing...
                            }

                        }
                        currentProgress++;
                    }
                }

                // - Teams based on XML templates
                var teams = hierarchy.Teams?.Teams;
                if (teams != null && teams.Any())
                {
                    foreach (var team in teams)
                    {
                        WriteSubProgress("Teams", "Team", currentProgress, totalCount);
                        if (PnPProvisioningContext.Current != null)
                        {
                            // Get a fresh Access Token for every request
                            accessToken = PnPProvisioningContext.Current.AcquireToken(GraphHelper.MicrosoftGraphBaseURI, "Group.ReadWrite.All");

                            if (accessToken != null)
                            {
                                // Create the Team starting from the XML PnP Provisioning Schema definition
                                CreateTeamFromProvisioningSchema(scope, parser, hierarchy.Connector, team, accessToken);
                            }
                        }

                        currentProgress++;
                    }
                }

                // - Apps
            }

            return parser;
        }

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ExtractConfiguration configuration)
        {
            using (var scope = new PnPMonitoredScope(Name))
            {
                var accessToken = PnPProvisioningContext.Current.AcquireTokenWithMultipleScopes(new Uri(GraphHelper.MicrosoftGraphBaseURI).Authority, "Group.ReadWrite.All", "User.Read.All");

                if (accessToken != null)
                {

                    if (configuration.Tenant.Teams.IncludeAllTeams)
                    {
                        // Retrieve all groups with teams

                        var groupIds = GetAllIdsForAllGroupsWithTeams(accessToken);
                        foreach (var groupId in groupIds)
                        {
                            Team team = ParseTeamJson(configuration, accessToken, groupId, scope);
                            if (team != null)
                            {
                                hierarchy.Teams.Teams.Add(team);
                            }
                        }
                    }
                    if (configuration.Tenant.Teams.TeamSiteUrls.Any())
                    {
                        foreach (var siteUrl in configuration.Tenant.Teams.TeamSiteUrls)
                        {
                            using (var siteContext = tenant.Context.Clone(siteUrl))
                            {
                                var groupId = siteContext.Web.GetPropertyBagValueString("GroupId", null);
                                if (groupId != null)
                                {
                                    Team team = ParseTeamJson(configuration, accessToken, groupId, scope);
                                    if (team != null)
                                    {
                                        hierarchy.Teams.Teams.Add(team);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (var siteUrl in PnPProvisioningContext.Current.ParsedSiteUrls)
                        {
                            using (var siteContext = tenant.Context.Clone(siteUrl))
                            {
                                var groupId = siteContext.Web.GetPropertyBagValueString("GroupId", null);
                                if (groupId != null)
                                {
                                    Team team = ParseTeamJson(configuration, accessToken, groupId, scope);
                                    if (team != null)
                                    {
                                        hierarchy.Teams.Teams.Add(team);
                                    }
                                }
                            }
                        }
                    }
                }
                return hierarchy;
            }
        }

        private static Team ParseTeamJson(ExtractConfiguration configuration, string accessToken, string groupId, PnPMonitoredScope scope)
        {
            var team = new Team();

            // Get Settings
            try
            {
                var teamString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{groupId}", accessToken);
                team = JsonConvert.DeserializeObject<Team>(teamString);
                if (configuration.Tenant.Teams.IncludeGroupId)
                {
                    team.GroupId = groupId;
                }
                team = GetTeamChannels(configuration, accessToken, groupId, team, scope);
                team = GetTeamApps(accessToken, groupId, team, scope);
                team = GetTeamSecurity(accessToken, groupId, team, scope);
                if (configuration.PersistAssetFiles)
                {
                    GetTeamPhoto(configuration, accessToken, groupId, team, scope);
                }
            }
            catch (ApplicationException ex)
            {
#if !NETSTANDARD2_0
                if (ex.InnerException is HttpException)
                {
                    if (((HttpException)ex.InnerException).GetHttpCode() == 404)
                    {
                        // no team, swallow
                    }
                    else
                    {
                        throw ex;
                    }
                }
                else
                {
                    throw ex;
                }
#else
                // untested change
                if (ex.Message.StartsWith("404"))
                {
                    // no team, swallow
                }
                else
                {
                    throw ex;
                }
#endif
            }
            return team;
        }

        private static void GetTeamPhoto(ExtractConfiguration configuration, string accessToken, string groupId, Team team, PnPMonitoredScope scope)
        {
            // get the photo stream
            var teamPhotoIdString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{groupId}/photo", accessToken);
            var teamPhotoId = JObject.Parse(teamPhotoIdString)["id"].Value<string>();
            var groupPhotoString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{groupId}/photos/{teamPhotoId}");
            var mediaType = JObject.Parse(groupPhotoString)["@odata.mediaContentType"].Value<string>();
            using (var teamPhotoStream = HttpHelper.MakeGetRequestForStream($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{groupId}/photos/{teamPhotoId}/$value", null, accessToken))
            {
                var extension = string.Empty;
                switch (mediaType)
                {
                    case "image/jpeg":
                        {
                            extension = ".jpg";
                            break;
                        }
                    case "image/gif":
                        {
                            extension = ".gif";
                            break;
                        }
                    case "image/png":
                        {
                            extension = ".png";
                            break;
                        }
                    case "image/bmp":
                        {
                            extension = ".bmp";
                            break;
                        }
                }
                configuration.FileConnector.SaveFileStream($"photo_{groupId}_{teamPhotoId}{extension}", $"TeamData/TEAM_{groupId}", teamPhotoStream);
                team.Photo = $"TeamData/TEAM_{groupId}/photo_{groupId}_{teamPhotoId}{extension}";
            }
        }

        private static Team GetTeamSecurity(string accessToken, string groupId, Team team, PnPMonitoredScope scope)
        {
            team.Security = new TeamSecurity();
            var teamOwnersString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{groupId}/owners?$select=userPrincipalName", accessToken);
            foreach (var user in JObject.Parse(teamOwnersString)["value"] as JArray)
            {
                team.Security.Owners.Add(user.ToObject<TeamSecurityUser>());
            }
            var teamMembersString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{groupId}/members?$select=userPrincipalName", accessToken);
            foreach (var user in JObject.Parse(teamMembersString)["value"] as JArray)
            {
                team.Security.Members.Add(user.ToObject<TeamSecurityUser>());
            }
            team.Security.AllowToAddGuests = GetAllowToAddGuestsSetting(scope, groupId, accessToken);

            return team;
        }

        private static Team GetTeamApps(string accessToken, string groupId, Team team, PnPMonitoredScope scope)
        {
            var teamsAppsString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{groupId}/installedApps", accessToken);
            foreach (var app in JObject.Parse(teamsAppsString)["value"] as JArray)
            {
                team.Apps.Add(new TeamAppInstance() { AppId = app["id"].Value<string>() });
            }
            return team;
        }

        private static Team GetTeamChannels(ExtractConfiguration configuration, string accessToken, string groupId, Team team, PnPMonitoredScope scope)
        {
            var teamChannelsString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{groupId}/channels", accessToken);
            team.Channels.AddRange(JsonConvert.DeserializeObject<List<Model.Teams.TeamChannel>>(JObject.Parse(teamChannelsString)["value"].ToString()));

            foreach (var channel in team.Channels)
            {
                channel.Tabs.AddRange(GetTeamChannelTabs(configuration, accessToken, groupId, channel.ID));
                if (configuration.Tenant.Teams.IncludeMessages)
                {
                    var channelMessagesString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{groupId}/channels/{channel.ID}/messages", accessToken);
                    foreach (var message in JObject.Parse(channelMessagesString)["value"] as JArray)
                    {
                        // We cannot set the createdDateTime value while posting messages, so remove it from the export.
                        var messageObject = CleanUpMessage((JObject)message);
                        channel.Messages.Add(new TeamChannelMessage() { Message = messageObject.ToString() });
                    }
                }
            }
            return team;
        }

        private static List<TeamTab> GetTeamChannelTabs(ExtractConfiguration configuration, string accessToken, string groupId, string channelId)
        {
            List<TeamTab> tabs = new List<TeamTab>();
            var teamTabsString = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/teams/{groupId}/channels/{channelId}/tabs", accessToken);
            foreach (var tab in JsonConvert.DeserializeObject<List<TeamTab>>(JObject.Parse(teamTabsString)["value"].ToString()))
            {
                if (tab.Configuration != null && string.IsNullOrEmpty(tab.Configuration.ContentUrl))
                {
                    tab.Configuration = null;
                }
                if (tab.Configuration != null)
                {
                    tab.Configuration.EntityId = tab.Configuration.EntityId ?? "";
                    tab.Configuration.WebsiteUrl = tab.Configuration.WebsiteUrl ?? "";
                    tab.Configuration.RemoveUrl = tab.Configuration.RemoveUrl ?? "";
                }
                tabs.Add(tab);
            }
            return tabs;
        }
        #endregion

        private static string CreateMailNicknameFromDisplayName(string displayName)
        {
            var mailNickname = displayName.ToLower();
            mailNickname = RemoveUnallowedCharacters(mailNickname);
            mailNickname = ReplaceAccentedCharactersWithLatin(mailNickname);
            return mailNickname;
        }

        private static string RemoveUnallowedCharacters(string str)
        {
            const string unallowedCharacters = "[&_,!@;:#¤`´~¨='%<>/\\\\\"\\.\\$\\*\\^\\+\\|\\{\\}\\[\\]\\-\\(\\)\\?\\s]";
            var regex = new Regex(unallowedCharacters);
            return regex.Replace(str, "");
        }

        private static string ReplaceAccentedCharactersWithLatin(string str)
        {
            const string a = "[äåàáâãæ]";
            var regex = new Regex(a, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "a");

            const string e = "[èéêëēĕėęě]";
            regex = new Regex(e, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "e");

            const string i = "[ìíîïĩīĭįı]";
            regex = new Regex(i, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "i");

            const string o = "[öòóôõø]";
            regex = new Regex(o, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "o");

            const string u = "[üùúû]";
            regex = new Regex(u, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "u");

            const string c = "[çċčćĉ]";
            regex = new Regex(c, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "c");

            const string d = "[ðďđđ]";
            regex = new Regex(d, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "d");

            return str;
        }
    }
}
#endif