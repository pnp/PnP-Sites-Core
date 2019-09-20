#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Graph;
using System;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
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
            String teamId = null;

            // If we have to Clone an existing Team
            if (!String.IsNullOrWhiteSpace(team.CloneFrom))
            {
                teamId = CloneTeam(scope, team, parser, accessToken);
            }
            // If we start from an already existing Group
            else if (!String.IsNullOrEmpty(team.GroupId))
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

            if (!String.IsNullOrEmpty(teamId))
            {
                // Wait to be sure that the Team is ready before configuring it
                WaitForTeamToBeReady(accessToken, teamId);

                // And now we configure security, channels, and apps
                // Only configure Security, if Security is configured
                if (team.Security != null) { 
                    if (!SetGroupSecurity(scope, team, teamId, accessToken)) return null;
                }
                if (!SetTeamChannels(scope, parser, team, teamId, accessToken)) return null;
                if (!SetTeamApps(scope, team, teamId, accessToken)) return null;

                // So far the Team's photo cannot be set if we don't have an already existing mailbox
                // if (!SetTeamPhoto(scope, parser, connector, team, teamId, accessToken)) return null;

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
            Boolean wait = true;
            Int32 iterations = 0;
            while (wait)
            {
                iterations++;

                try
                {
                    var jsonOwners = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/owners?$select=id", accessToken);
                    if (!String.IsNullOrEmpty(jsonOwners))
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

        /// <summary>
        /// Checks if a Group exists by ID
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="groupId">The ID of the Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Group exists or not</returns>
        private static Boolean GroupExistsById(PnPMonitoredScope scope, string groupId, string accessToken)
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
        private static String GetGroupIdByMailNickname(PnPMonitoredScope scope, string mailNickname, string accessToken)
        {
            var alreadyExistingGroupId = !String.IsNullOrEmpty(mailNickname) ?
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
            if (String.IsNullOrEmpty(alreadyExistingGroupId))
            {
                // Otherwise we create the Group, first

                // Prepare the IDs for owners and members
                String[] desideredOwnerIds;
                String[] desideredMemberIds;
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

                    desideredOwnerIds = team.Security.Owners.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
                    desideredMemberIds = team.Security.Members.Select(o => userIdsByUPN[o.UserPrincipalName]).Union(desideredOwnerIds).ToArray();
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingUserError, ex.Message);
                    return (null);
                }

                var groupCreationRequest = new
                {
                    displayName = parser.ParseString(team.DisplayName),
                    description = parser.ParseString(team.Description),
                    groupTypes = new String[]
                    {
                        "Unified"
                    },
                    mailEnabled = true,
                    mailNickname = parsedMailNickname,
                    securityEnabled = false,
                    visibility = team.Visibility.ToString(),
                    owners_odata_bind = (from o in desideredOwnerIds select $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{Uri.EscapeDataString(o.Replace("'", "''"))}").ToArray(),
                    members_odata_bind = (from m in desideredMemberIds select $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{Uri.EscapeDataString(m.Replace("'", "''"))}").ToArray()
                };

                // Make the Graph request to create the Office 365 Group
                var createdGroupJson = HttpHelper.MakePostRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups",
                    groupCreationRequest, HttpHelper.JsonContentType, accessToken);
                var createdGroupId = JToken.Parse(createdGroupJson).Value<string>("id");

                // Wait for the Group to be ready
                Boolean wait = true;
                Int32 iterations = 0;
                while (wait)
                {
                    iterations++;

                    try
                    {
                        var jsonGroup = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{createdGroupId}", accessToken);
                        if (!String.IsNullOrEmpty(jsonGroup))
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
            var teamId = CreateOrUpdateTeamFromGroup(scope, team, parser, team.GroupId, accessToken);

            return (teamId);
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
        private static Object PrepareTeamCloneRequestContent(Team team, TokenParser parser)
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
        private static string CreateOrUpdateTeamFromGroup(PnPMonitoredScope scope, Team team, TokenParser parser, String groupId, string accessToken)
        {
            var content = PrepareTeamRequestContent(team, parser);

            var teamId = GraphHelper.CreateOrUpdateGraphObject(scope,
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

            return (teamId);
        }

        /// <summary>
        /// Prepares the JSON object for the request to create/update a Team
        /// </summary>
        /// <param name="team">The Domain Model Team object</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <returns>The JSON object ready to be serialized into the JSON request</returns>
        private static Object PrepareTeamRequestContent(Team team, TokenParser parser)
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
                    team.MemberSettings?.AllowCreateUpdateRemoveConnectors
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
        private static void ArchiveTeam(PnPMonitoredScope scope, String teamId, Boolean archived, String accessToken)
        {
            try
            {
                if (archived)
                {
                    // Archive the Team
                    HttpHelper.MakePostRequest(
                        $"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/archive", accessToken: accessToken);
                }
                else
                {
                    // Unarchive the Team
                    HttpHelper.MakePostRequest(
                        $"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/unarchive", accessToken: accessToken);
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
        /// <param name="team">The Team settings, including security settings</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Security settings have been provisioned or not</returns>
        private static bool SetGroupSecurity(PnPMonitoredScope scope, Team team, string teamId, string accessToken)
        {
            String[] desideredOwnerIds;
            String[] desideredMemberIds;
            String[] finalOwnerIds;
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

                desideredOwnerIds = team.Security.Owners.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
                desideredMemberIds = team.Security.Members.Select(o => userIdsByUPN[o.UserPrincipalName]).Union(desideredOwnerIds).ToArray();
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingUserError, ex.Message);
                return false;
            }

            String[] ownerIdsToAdd;
            String[] ownerIdsToRemove;
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

            String[] memberIdsToAdd;
            String[] memberIdsToRemove;
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
            return JToken.Parse(HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels", accessToken))["value"];
        }

        private static string UpdateTeamChannel(TeamChannel channel, string teamId, JToken existingChannel, string accessToken)
        {
            // Not supported to update 'General' Channel
            if(channel.DisplayName.Equals("General", StringComparison.InvariantCultureIgnoreCase))
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

        private static string CreateTeamChannel(PnPMonitoredScope scope, TeamChannel channel, string teamId, string accessToken)
        {
            var channelToCreate = new
            {
                channel.Description,
                channel.DisplayName,
                channel.IsFavoriteByDefault
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

                var existingTab = existingTabs.FirstOrDefault(x => HttpUtility.UrlDecode(x["displayName"].ToString()) == tab.DisplayName && x["teamsAppId"].ToString() == tab.TeamsAppId);

                var tabId = existingTab == null ? CreateTeamTab(scope, tab, parser, teamId, channelId, accessToken) : UpdateTeamTab(tab, parser, teamId, channelId, existingTab["id"].ToString(), accessToken);

                if (tabId == null) return false;
            }

            return true;
        }

        public static JToken GetExistingTeamChannelTabs(string teamId, string channelId, string accessToken)
        {
            return JToken.Parse(HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/tabs", accessToken))["value"];
        }

        private static string UpdateTeamTab(TeamTab tab, TokenParser parser, string teamId, string channelId, string tabId, string accessToken)
        {
            var displayname = parser.ParseString(tab.DisplayName);

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

            HttpHelper.MakePatchRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/tabs/{tabId}", tabToUpdate, HttpHelper.JsonContentType, accessToken);

            // Add the teamsAppId back now that we've updated the tab
            tab.TeamsAppId = teamsAppId;

            return tabId;
        }

        private static string CreateTeamTab(PnPMonitoredScope scope, TeamTab tab, TokenParser parser, string teamId, string channelId, string accessToken)
        {
            var displayname = parser.ParseString(tab.DisplayName);
            var teamsAppId = parser.ParseString(tab.TeamsAppId);

            if (tab.Configuration != null)
            {
                tab.Configuration.EntityId = parser.ParseString(tab.Configuration.EntityId);
                tab.Configuration.ContentUrl = parser.ParseString(tab.Configuration.ContentUrl);
                tab.Configuration.RemoveUrl = parser.ParseString(tab.Configuration.RemoveUrl);
                tab.Configuration.WebsiteUrl = parser.ParseString(tab.Configuration.WebsiteUrl);
            }

            var tabToCreate = new
            {
                displayname,
                teamsAppId,
                configuration = tab.Configuration != null
                    ? new
                    {
                        tab.Configuration.EntityId,
                        tab.Configuration.ContentUrl,
                        tab.Configuration.RemoveUrl,
                        tab.Configuration.WebsiteUrl
                    }
                    : null
            };

            var tabId = GraphHelper.CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST,
                $"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/tabs",
                tabToCreate,
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

        private static string CreateTeamChannelMessage(PnPMonitoredScope scope, TokenParser parser, TeamChannelMessage message, string teamId, string channelId, string accessToken)
        {
            var messageString = parser.ParseString(message.Message);
            var messageJson = default(JToken);

            try
            {
                // If the message is already in JSON format, we just use it
                messageJson = JToken.Parse(messageString);
            }
            catch
            {
                // Otherwise try to build the JSON message content from scratch
                messageJson = JToken.Parse($"{{ \"body\": {{ \"content\": \"{messageString}\" }} }}");
            }

            var messageId = GraphHelper.CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST,
                $"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/messages",
                messageJson,
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
        /// <param name="team">The Team settings, including security settings</param>
        /// <param name="teamId">The ID of the target Team</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Apps have been provisioned or not</returns>
        private static bool SetTeamApps(PnPMonitoredScope scope, Team team, string teamId, string accessToken)
        {
            foreach (var app in team.Apps)
            {
                Object appToCreate = new JObject
                {
                    ["teamsApp@odata.bind"] = app.AppId
                };

                var id = GraphHelper.CreateOrUpdateGraphObject(scope,
                    HttpMethodVerb.POST,
                    $"{GraphHelper.MicrosoftGraphBaseURI}beta/teams/{teamId}/installedApps",
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
            if (!String.IsNullOrEmpty(team.Photo) && connector != null)
            {
                var photoPath = parser.ParseString(team.Photo);
                var photoBytes = ConnectorFileHelper.GetFileBytes(connector, team.Photo);

                using (var mem = new MemoryStream())
                {
                    mem.Write(photoBytes, 0, photoBytes.Length);
                    mem.Position = 0;

                    HttpHelper.MakePostRequest(
                        $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}/photo/$value",
                        mem, "image/jpeg", accessToken);
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

            try
            {
                var teamId = responseHeaders.Location.ToString().Split('\'')[1];
                var team = HttpHelper.MakeGetRequestForString($"{GraphHelper.MicrosoftGraphBaseURI}v1.0/groups/{teamId}", accessToken);
                return JToken.Parse(team);
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_FetchingError, ex.Message);
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
            if (teamTemplate.Classification != null) team["classification"] = teamTemplate.Classification;
            team["visibility"] = teamTemplate.Visibility.ToString();

            return team.ToString();
        }

        #region PnP Provisioning Engine infrastructural code

        public override bool WillProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = hierarchy.Teams?.TeamTemplates?.Any() |
                    hierarchy.Teams?.Teams?.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }

        public override TokenParser ProvisionObjects(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(Name))
            {
                // Prepare a method global variable to store the Access Token
                String accessToken = null;

                // - Teams based on JSON templates
                var teamTemplates = hierarchy.Teams?.TeamTemplates;
                if (teamTemplates != null && teamTemplates.Any())
                {
                    foreach (var teamTemplate in teamTemplates)
                    {
                        if (PnPProvisioningContext.Current != null)
                        {
                            // Get a fresh Access Token for every request
                            accessToken = PnPProvisioningContext.Current.AcquireToken(GraphHelper.MicrosoftGraphBaseURI, "Group.ReadWrite.All");

                            if (accessToken != null)
                            {
                                // Create the Team starting from the JSON template
                                var team = CreateTeamFromJsonTemplate(scope, parser, teamTemplate, accessToken);

                                // TODO: possible further processing...
                            }

                        }

                    }
                }

                // - Teams based on XML templates
                var teams = hierarchy.Teams?.Teams;
                if (teams != null && teams.Any())
                {
                    foreach (var team in teams)
                    {
                        // Get a fresh Access Token for every request
                        accessToken = PnPProvisioningContext.Current.AcquireToken(GraphHelper.MicrosoftGraphBaseURI, "Group.ReadWrite.All");

                        // Create the Team starting from the XML PnP Provisioning Schema definition
                        CreateTeamFromProvisioningSchema(scope, parser, hierarchy.Connector, team, accessToken);

                        // TODO: possible further processing...
                    }
                }

                // - Apps
            }

            return parser;
        }

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ProvisioningTemplateCreationInformation creationInfo)
        {
            // So far, no extraction
            return hierarchy;
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
            const string a = "[äåàáâã]";
            var regex = new Regex(a, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "a");

            const string e = "[èéêë]";
            regex = new Regex(e, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "e");

            const string i = "[ìíîï]";
            regex = new Regex(i, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "i");

            const string o = "[öòóôõ]";
            regex = new Regex(o, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "o");

            const string u = "[üùúû]";
            regex = new Regex(u, RegexOptions.IgnoreCase);
            str = regex.Replace(str, "u");

            return str;
        }
    }
}
#endif