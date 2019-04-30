#if !ONPREMISES
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Linq;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Utilities;
using System.Net;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Web;
using System.Net.Http;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Object Handler to manage Microsoft Teams stuff
    /// </summary>
    internal class ObjectTeams : ObjectHierarchyHandlerBase
    {
        private static String jsonContentType = "application/json";

        public override string Name => "Teams";

        /// <summary>
        /// Creates a new Team from a PnP Provisioning Schema definition
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="team">The Team to provision</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The provisioned Team as a JSON object</returns>
        private static JToken CreateTeamFromProvisioningSchema(PnPMonitoredScope scope, TokenParser parser, Team team, string accessToken)
        {
            String teamId = null;

            // If we have to Clone an existing Team
            if (!String.IsNullOrWhiteSpace(team.CloneFrom))
            {
                // TODO: handle cloning
                scope.LogError("Cloning not supported yet");
                return null;
            }
            // If we start from an already existing Group
            else if (!String.IsNullOrEmpty(team.GroupId))
            {
                // Check if the Group exists
                if (GroupExists(scope, team.GroupId, accessToken))
                {
                    // Then promote the Group into a Team or update it, if it already exists
                    teamId = CreateOrUpdateTeamFromGroup(scope, team, accessToken);
                }
                else
                {
                    // Log the exception and return NULL (i.e. cancel)
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_GroupDoesNotExists, team.GroupId);
                    return null;
                }
            }
            // Otherwise create a Team from scratch
            else
            {
                teamId = CreateOrUpdateTeam(scope, team, accessToken);
            }

            if (!String.IsNullOrEmpty(teamId))
            {
                // Wait for the Team to be ready
                Boolean wait = true;
                Int32 iterations = 0;
                while (wait)
                {
                    iterations++;

                    try
                    {
                        var jsonOwners = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/groups/{teamId}/owners?$select=id", accessToken);
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

                // And now we configure security, channels, and apps
                if (!SetGroupSecurity(scope, team, teamId, accessToken)) return null;
                if (!SetTeamChannels(scope, parser, team, teamId, accessToken)) return null;
                if (!SetTeamApps(scope, team, teamId, accessToken)) return null;

                // Call Archive or Unarchive for the current Team
                ArchiveTeam(scope, teamId, team.Archived, accessToken);

                try
                {
                    // Get the whole Team that we just created and return it back as the method result
                    return JToken.Parse(HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/teams/{teamId}", accessToken));
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingError, ex.Message);
                }
            }

            return null;
        }

        /// <summary>
        /// Checks if a Group exists
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="groupId">The ID of the Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Group exists or not</returns>
        private static Boolean GroupExists(PnPMonitoredScope scope, string groupId, string accessToken)
        {
            try
            {
                var existingGroup = JToken.Parse(HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/groups/{groupId}", accessToken));
                if (existingGroup.Value<String>("id") == groupId)
                {
                    return (true);
                }
                else
                {
                    return (false);
                }
            }
            catch (Exception)
            {
                return (false);
            }

        }

        /// <summary>
        /// Creates or updates a Team object via Graph
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="team">The Team to create</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The ID of the created or update Team</returns>
        private static string CreateOrUpdateTeam(PnPMonitoredScope scope, Team team, string accessToken)
        {
            var content = PrepareTeamRequestContent(team);

            var teamId = CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST_WITH_RESPONSE_HEADERS,
                $"https://graph.microsoft.com/beta/teams",
                content,
                jsonContentType,
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
        /// Creates or updates a Team object via Graph promoting an existing Group
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="team">The Team to create</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The ID of the created or updated Team</returns>
        private static string CreateOrUpdateTeamFromGroup(PnPMonitoredScope scope, Team team, string accessToken)
        {
            var content = PrepareTeamRequestContent(team);

            var teamId = CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST,
                $"https://graph.microsoft.com/beta/groups/{team.GroupId}/team",
                content,
                jsonContentType,
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
        /// Prepares the JSON object for the request to create/update a Team
        /// </summary>
        /// <param name="team">The Domain Model Team object</param>
        /// <returns>The JSON object ready to be serialized into the JSON request</returns>
        private static Object PrepareTeamRequestContent(Team team)
        {
            var content = new
            {
                template_odata_bind = "https://graph.microsoft.com/beta/teamsTemplates('standard')",
                team.DisplayName,
                team.Description,
                //team.Classification,
                //team.Specialization,
                //team.Visibility,
                funSettings = new
                {
                    team.FunSettings.AllowGiphy,
                    team.FunSettings.GiphyContentRating,
                    team.FunSettings.AllowStickersAndMemes,
                    team.FunSettings.AllowCustomMemes,
                },
                guestSettings = new
                {
                    team.GuestSettings.AllowCreateUpdateChannels,
                    team.GuestSettings.AllowDeleteChannels,
                },
                memberSettings = new
                {
                    team.MemberSettings.AllowCreateUpdateChannels,
                    team.MemberSettings.AllowAddRemoveApps,
                    team.MemberSettings.AllowDeleteChannels,
                    team.MemberSettings.AllowCreateUpdateRemoveTabs,
                    team.MemberSettings.AllowCreateUpdateRemoveConnectors
                },
                messagingSettings = new
                {
                    team.MessagingSettings.AllowUserEditMessages,
                    team.MessagingSettings.AllowUserDeleteMessages,
                    team.MessagingSettings.AllowOwnerDeleteMessages,
                    team.MessagingSettings.AllowTeamMentions,
                    team.MessagingSettings.AllowChannelMentions
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
                        $"https://graph.microsoft.com/beta/teams/{teamId}/archive", accessToken: accessToken);
                }
                else
                {
                    // Unarchive the Team
                    HttpHelper.MakePostRequest(
                        $"https://graph.microsoft.com/beta/teams/{teamId}/unarchive", accessToken: accessToken);
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
            try
            {
                var userIdsByUPN = team.Security.Owners
                    .Select(o => o.UserPrincipalName)
                    .Concat(team.Security.Members.Select(m => m.UserPrincipalName))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(k => k, k =>
                    {
                        var jsonUser = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/users/{k}?$select=id", accessToken);
                        return JToken.Parse(jsonUser).Value<string>("id");
                    });

                desideredOwnerIds = team.Security.Owners.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
                desideredMemberIds = team.Security.Members.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
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
                var jsonOwners = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/groups/{teamId}/owners?$select=id", accessToken);

                string[] currentOwnerIds = GetIdsFromList(jsonOwners);

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
                        ["@odata.id"] = $"https://graph.microsoft.com/v1.0/users/{ownerId}"
                    };
                    HttpHelper.MakePostRequest($"https://graph.microsoft.com/v1.0/groups/{teamId}/owners/$ref", content, "application/json", accessToken);
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
                    HttpHelper.MakeDeleteRequest($"https://graph.microsoft.com/v1.0/groups/{teamId}/owners/{ownerId}/$ref", accessToken);
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
                var jsonOwners = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/groups/{teamId}/members?$select=id", accessToken);

                string[] currentMemberIds = GetIdsFromList(jsonOwners);

                // Exclude members already into the group
                memberIdsToAdd = desideredMemberIds.Except(currentMemberIds).ToArray();

                if (team.Security.ClearExistingMembers)
                {
                    memberIdsToRemove = currentMemberIds.Except(desideredMemberIds).ToArray();
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
                        ["@odata.id"] = $"https://graph.microsoft.com/v1.0/users/{memberId}"
                    };
                    HttpHelper.MakePostRequest($"https://graph.microsoft.com/v1.0/groups/{teamId}/members/$ref", content, "application/json", accessToken);
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
                    HttpHelper.MakeDeleteRequest($"https://graph.microsoft.com/v1.0/groups/{teamId}/members/{memberId}/$ref", accessToken);
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
            // TODO: create resource strings for exceptions

            if (team.Channels != null)
            {
                foreach (var channel in team.Channels)
                {
                    // Create the channel object for the API call
                    var channelToCreate = new
                    {
                        channel.Description,
                        channel.DisplayName,
                        channel.IsFavoriteByDefault
                    };

                    var channelId = CreateOrUpdateGraphObject(scope,
                        HttpMethodVerb.POST,
                        $"https://graph.microsoft.com/beta/teams/{teamId}/channels",
                        channelToCreate,
                        jsonContentType,
                        accessToken,
                        "NameAlreadyExists",
                        CoreResources.Provisioning_ObjectHandlers_Teams_Team_ChannelAlreadyExists,
                        "displayName",
                        channel.DisplayName,
                        CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError,
                        canPatch: false);

                    if (channelId == null) return false;

                    // If there are any Tabs for the current channel
                    if (channel.Tabs == null || !channel.Tabs.Any()) continue;

                    foreach (var tab in channel.Tabs)
                    {
                        // Create the object for the API call
                        var tabToCreate = new
                        {
                            tab.DisplayName,
                            tab.TeamsAppId,
                            configuration = tab.Configuration != null ? new
                            {
                                tab.Configuration.EntityId,
                                tab.Configuration.ContentUrl,
                                tab.Configuration.RemoveUrl,
                                tab.Configuration.WebsiteUrl
                            } : null
                        };

                        var tabId = CreateOrUpdateGraphObject(scope,
                            HttpMethodVerb.POST,
                            $"https://graph.microsoft.com/beta/teams/{teamId}/channels/{channelId}/tabs",
                            tabToCreate,
                            jsonContentType,
                            accessToken,
                            "NameAlreadyExists",
                            CoreResources.Provisioning_ObjectHandlers_Teams_Team_TabAlreadyExists,
                            "displayName",
                            tab.DisplayName,
                            CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError,
                            canPatch: false);

                        if (tabId == null) return false;
                    }

                    // TODO: Handle TabResources
                    // We need to define a "schema" for their settings

                    // If there are any messages for the current channel
                    if (channel.Messages == null || !channel.Messages.Any()) continue;

                    foreach (var message in channel.Messages)
                    {
                        // Get and parse the CData
                        var messageString = parser.ParseString(message.Message);
                        var messageJson = JToken.Parse(messageString);

                        var messageId = CreateOrUpdateGraphObject(scope,
                            HttpMethodVerb.POST,
                            $"https://graph.microsoft.com/beta/teams/{teamId}/channels/{channelId}/messages",
                            messageJson,
                            jsonContentType,
                            accessToken,
                            null,
                            null,
                            null,
                            null,
                            CoreResources.Provisioning_ObjectHandlers_Teams_Team_CannotSendMessage,
                            canPatch: false);

                        if (messageId == null) return false;
                    }
                }
            }

            return true;
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

                var id = CreateOrUpdateGraphObject(scope,
                    HttpMethodVerb.POST,
                    $"https://graph.microsoft.com/beta/teams/{teamId}/installedApps",
                    appToCreate,
                    jsonContentType,
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
                responseHeaders = HttpHelper.MakePostRequestForHeaders("https://graph.microsoft.com/beta/teams", content, "application/json", accessToken);
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_ProvisioningError, ex.Message);
                return null;
            }

            try
            {
                var teamId = responseHeaders.Location.ToString().Split('\'')[1];
                var team = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/groups/{teamId}", accessToken);
                return JToken.Parse(team);
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_FetchingError, ex.Message);
            }

            return null;
        }

        /// <summary>
        /// Retrieves the IDs of items in a JSON list
        /// </summary>
        /// <param name="json">The JSON list</param>
        /// <returns>The array of IDs</returns>
        private static string[] GetIdsFromList(string json)
        {
            return JsonConvert.DeserializeAnonymousType(json, new { value = new[] { new { id = "" } } }).value.Select(v => v.id).ToArray();
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
            // if (teamTemplate.Classification != null) team["classification"] = teamTemplate.Classification;
            // team["visibility"] = teamTemplate.Visibility.ToString();

            return team.ToString();
        }

        /// <summary>
        /// Helper method to create or update an object through the Microsoft Graph
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="method">The HTTP method to use</param>
        /// <param name="uri">The URI for the Graph request</param>
        /// <param name="content">The content of the Graph request</param>
        /// <param name="contentType">The content type of the Graph request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <param name="alreadyExistsErrorMessage">The error message token that identifies an already existing item</param>
        /// <param name="warningMessage">The warning message to log when the target item already exists</param>
        /// <param name="matchingFieldName">The name of a field to match an already existing instance of the target item</param>
        /// <param name="matchingFieldValue">The value of a field to match an already existing instance of the target item</param>
        /// <param name="errorMessage">The error message to log when the create or update action fails</param>
        /// <param name="canPatch">Defines whether a Patch HTTP request can be executed to update an already existing target item</param>
        /// <returns>The ID of the create or updated target item</returns>
        private static String CreateOrUpdateGraphObject(
            PnPMonitoredScope scope,
            HttpMethodVerb method,
            String uri,
            Object content,
            String contentType,
            String accessToken,
            String alreadyExistsErrorMessage,
            String warningMessage,
            String matchingFieldName,
            String matchingFieldValue,
            String errorMessage,
            Boolean canPatch
            )
        {
            try
            {
                String itemId = null;
                String json = null;
                HttpResponseHeaders responseHeaders;

                // Try to create the Graph object
                switch (method)
                {
                    case HttpMethodVerb.POST:
                        json = HttpHelper.MakePostRequestForString(uri, content, contentType, accessToken);
                        itemId = JToken.Parse(json).Value<String>("id");
                        break;
                    case HttpMethodVerb.PUT:
                        json = HttpHelper.MakePutRequestForString(uri, content, contentType, accessToken);
                        itemId = JToken.Parse(json).Value<String>("id");
                        break;
                    case HttpMethodVerb.POST_WITH_RESPONSE_HEADERS:
                        responseHeaders = HttpHelper.MakePostRequestForHeaders(uri, content, contentType, accessToken);
                        itemId = responseHeaders.Location.ToString().Split('\'')[1];
                        break;
                }

                // Return the ID of the just created item
                return itemId;
            }
            catch (Exception ex)
            {
                // In case of exception, let's see if the target item already exists
                if (!String.IsNullOrEmpty(alreadyExistsErrorMessage) && 
                    !String.IsNullOrEmpty(matchingFieldName) &&
                    !String.IsNullOrEmpty(matchingFieldValue) &&
                    ex.InnerException.Message.Contains(alreadyExistsErrorMessage))
                {
                    try
                    {
                        if (!String.IsNullOrEmpty(warningMessage))
                        {
                            scope.LogWarning(warningMessage);
                        }

                        // If it's a POST we need to look for any existing item
                        String id = null;

                        // In case of PUT we already have the id
                        if (method == HttpMethodVerb.POST)
                        {
                            // Filter by field and value specified
                            String json = HttpHelper.MakeGetRequestForString($"{uri}?$select=id&$filter={matchingFieldName}%20eq%20'{WebUtility.UrlEncode(matchingFieldValue)}'");
                            // Get the id of existing item
                            id = GetIdsFromList(json)[0];
                            uri = $"{uri}/{id}";
                        }

                        // Patch the item, if supported
                        if (canPatch)
                        {
                            HttpHelper.MakePatchRequestForString(uri, content, contentType, accessToken);
                        }

                        return id;
                    }
                    catch (Exception exUpdate)
                    {
                        if (!String.IsNullOrEmpty(errorMessage))
                        {
                            scope.LogError(errorMessage, exUpdate.Message);
                        }
                        return null;
                    }
                }
                else
                {
                    return (null);
                }
            }
        }

#region PnP Provisioning Engine infrastructural code

        public override bool WillProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#if !ONPREMISES
            if (!_willProvision.HasValue)
            {
                _willProvision = hierarchy.Teams?.TeamTemplates?.Any() |
                    hierarchy.Teams?.Teams?.Any();
            }
#else
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
#endif
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
#if !ONPREMISES
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
                        // Get a fresh Access Token for every request
                        accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.ReadWrite.All");

                        // Create the Team starting from the JSON template
                        var team = CreateTeamFromJsonTemplate(scope, parser, teamTemplate, accessToken);

                        // TODO: possible further processing...
                    }
                }

                // - Teams based on XML templates
                var teams = hierarchy.Teams?.Teams;
                if (teams != null && teams.Any())
                {
                    foreach (var team in teams)
                    {
                        // Get a fresh Access Token for every request
                        accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.ReadWrite.All");

                        // Create the Team starting from the XML PnP Provisioning Schema definition
                        CreateTeamFromProvisioningSchema(scope, parser, team, accessToken);

                        // TODO: possible further processing...
                    }
                }

                // - Apps
            }
#endif

            return parser;
        }

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ProvisioningTemplateCreationInformation creationInfo)
        {
            // So far, no extraction
            return hierarchy;
        }

#endregion
    }

    enum HttpMethodVerb
    {
        GET,
        POST,
        PUT,
        PATCH,
        POST_WITH_RESPONSE_HEADERS
    }
}
#endif