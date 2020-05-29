#if !ONPREMISES
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectTeamsTests
    {
        public const String MicrosoftGraphBaseURI = "https://graph.microsoft.com/";
        private readonly List<string> _teamNames = new List<string>();
        private string _jsonTemplate;
        private Team _team;

        #region Init and Cleanup code

        [TestInitialize]
        public void Initialize()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                const string teamTemplateName = "Sample Engineering Team";
                _teamNames.Add(teamTemplateName);
                _jsonTemplate = "{ \"template@odata.bind\": \"https://graph.microsoft.com/beta/teamsTemplates(\'standard\')\", \"visibility\": \"Private\", \"displayName\": \"" + teamTemplateName + "\", \"description\": \"This is a sample engineering team, used to showcase the range of properties supported by this API\", \"channels\": [ { \"displayName\": \"Announcements 📢\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements.\" }, { \"displayName\": \"Training 🏋️\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample training channel, that is favorited by default, and contains an example of pinned website and YouTube tabs.\", \"tabs\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.web\')\", \"name\": \"A Pinned Website\", \"configuration\": { \"contentUrl\": \"https://docs.microsoft.com/en-us/microsoftteams/microsoft-teams\" } }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.youtube\')\", \"name\": \"A Pinned YouTube Video\", \"configuration\": { \"contentUrl\": \"https://tabs.teams.microsoft.com/Youtube/Home/YoutubeTab?videoId=X8krAMdGvCQ\", \"websiteUrl\": \"https://www.youtube.com/watch?v=X8krAMdGvCQ\" } } ] }, { \"displayName\": \"Planning 📅 \", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\", \"isFavoriteByDefault\": false }, { \"displayName\": \"Issues and Feedback 🐞\", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\" } ], \"memberSettings\": { \"allowCreateUpdateChannels\": true, \"allowDeleteChannels\": true, \"allowAddRemoveApps\": true, \"allowCreateUpdateRemoveTabs\": true, \"allowCreateUpdateRemoveConnectors\": true }, \"guestSettings\": { \"allowCreateUpdateChannels\": false, \"allowDeleteChannels\": false }, \"funSettings\": { \"allowGiphy\": true, \"giphyContentRating\": \"Moderate\", \"allowStickersAndMemes\": true, \"allowCustomMemes\": true }, \"messagingSettings\": { \"allowUserEditMessages\": true, \"allowUserDeleteMessages\": true, \"allowOwnerDeleteMessages\": true, \"allowTeamMentions\": true, \"allowChannelMentions\": true }, \"installedApps\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.vsts\')\" }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'1542629c-01b3-4a6d-8f76-1938b779e48d\')\" } ] }";

                const string teamName = @"Unallowed &_,.;:/\""!@$%^*[]+=&'|<>{}-()?#¤`´~¨ Ääkköset";
                _teamNames.Add(teamName);
                var security = new TeamSecurity
                {
                    AllowToAddGuests = false,
                    Owners = { new TeamSecurityUser { UserPrincipalName = ConfigurationManager.AppSettings["SPOUserName"] } }
                };
                var funSettings = new TeamFunSettings
                {
                    AllowCustomMemes = true,
                    AllowGiphy = true,
                    AllowStickersAndMemes = true,
                    GiphyContentRating = TeamGiphyContentRating.Moderate
                };
                var guestSettings = new TeamGuestSettings
                {
                    AllowCreateUpdateChannels = false,
                    AllowDeleteChannels = false
                };
                var memberSettings = new TeamMemberSettings
                {
                    AllowDeleteChannels = false,
                    AllowAddRemoveApps = true,
                    AllowCreateUpdateChannels = true,
                    AllowCreateUpdateRemoveConnectors = false,
                    AllowCreateUpdateRemoveTabs = true
                };
                var messagingSettings = new TeamMessagingSettings
                {
                    AllowChannelMentions = true,
                    AllowOwnerDeleteMessages = true,
                    AllowTeamMentions = false,
                    AllowUserDeleteMessages = true,
                    AllowUserEditMessages = true
                };
                var channel = new TeamChannel
                {
                    DisplayName = "Another channel",
                    Description = "Another channel description!",
                    IsFavoriteByDefault = true
                };
                var tab = new TeamTab
                {
                    DisplayName = "OneNote Tab",
                    TeamsAppId = "0d820ecd-def2-4297-adad-78056cde7c78"/*,
                Configuration = new TeamTabConfiguration
                {
                    ContentUrl = "https://todo-improve-this-test-when-tab-resource-provisioning-has-been-implemented"
                }*/
                };
                channel.Tabs.Add(tab);
                var message = new TeamChannelMessage
                {
                    Message = "Welcome to this awesome new channel!"
                };
                channel.Messages.Add(message);

                _team = new Team { DisplayName = teamName, Description = "Testing creating mailNickname from a display name that has unallowed and accented characters", Visibility = TeamVisibility.Public, Security = security, FunSettings = funSettings, GuestSettings = guestSettings, MemberSettings = memberSettings, MessagingSettings = messagingSettings, Channels = { channel } };
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
                {
                    foreach (var teamName in _teamNames)
                    {
                        var teams = GetTeamsByDisplayName(teamName);

                        foreach (var team in teams)
                        {
                            try
                            {
                                DeleteTeam(team["id"].ToString());
                            }
                            catch (Exception ex)
                            {
                                // NOOP
                            }
                        }
                    }
                }
            }
        }

        #endregion

        #region Private helper methods

        private static JToken GetTeamsByDisplayName(string displayName)
        {
            if (String.IsNullOrEmpty(displayName)) return null;

            var accessToken = PnPProvisioningContext.Current.AcquireToken(MicrosoftGraphBaseURI, "Group.Read.All");
            var requestUrl = $"{MicrosoftGraphBaseURI}v1.0/groups?$filter=displayName eq '{HttpUtility.UrlEncode(displayName.Replace("'", "''"))}'";
            return JToken.Parse(HttpHelper.MakeGetRequestForString(requestUrl, accessToken))["value"];
        }

        private static JToken GetTeamById(string teamId)
        {
            if (String.IsNullOrEmpty(teamId)) return null;

            var accessToken = PnPProvisioningContext.Current.AcquireToken(MicrosoftGraphBaseURI, "Group.Read.All");
            return JToken.Parse(HttpHelper.MakeGetRequestForString($"{MicrosoftGraphBaseURI}beta/teams/{teamId}", accessToken));
        }

        private static JToken GetTeamChannels(string teamId)
        {
            if (String.IsNullOrEmpty(teamId)) return null;

            var accessToken = PnPProvisioningContext.Current.AcquireToken(MicrosoftGraphBaseURI, "Group.Read.All");
            return JToken.Parse(HttpHelper.MakeGetRequestForString($"{MicrosoftGraphBaseURI}beta/teams/{teamId}/channels", accessToken))["value"];
        }

        private static JToken GetTeamChannelTabs(string teamId, string channelId)
        {
            if (String.IsNullOrEmpty(teamId)) return null;
            if (String.IsNullOrEmpty(channelId)) return null;

            var accessToken = PnPProvisioningContext.Current.AcquireToken(MicrosoftGraphBaseURI, "Group.Read.All");
            return JToken.Parse(HttpHelper.MakeGetRequestForString($"{MicrosoftGraphBaseURI}beta/teams/{teamId}/channels/{channelId}/tabs", accessToken))["value"];
        }

        private static bool GetAllowToAddGuests(string teamId) {
            var accessToken = PnPProvisioningContext.Current.AcquireToken(MicrosoftGraphBaseURI, "Group.Read.All");
            var response = JToken.Parse(HttpHelper.MakeGetRequestForString($"{MicrosoftGraphBaseURI}v1.0/groups/{teamId}/settings", accessToken));
            var groupGuestSettings = response["value"]?.FirstOrDefault(x => x["templateId"].ToString() == "08d542b9-071f-4e16-94b0-74abb372e3d9");
            return (bool)groupGuestSettings["values"]?.FirstOrDefault(x => x["name"].ToString() == "AllowToAddGuests")["value"];
        }

        private static void DeleteTeam(string id)
        {
            var accessToken = PnPProvisioningContext.Current.AcquireToken(MicrosoftGraphBaseURI, "Group.ReadWrite.All");

            var requestUrl = $"{MicrosoftGraphBaseURI}v1.0/groups/{id}";
            HttpHelper.MakeDeleteRequest(requestUrl, accessToken);

            //var accessToken = PnPProvisioningContext.Current.AcquireToken("https://api.spaces.skype.com", "user_impersonation");

            //var requestUrl = $"https://teams.microsoft.com/api/mt/emea/beta/teams/{id}/delete";
            //HttpHelper.MakeDeleteRequest(requestUrl, accessToken);
        }

        private static void Provision(ProvisioningTemplate template)
        {
            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                using (var ctx = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(ctx);
                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectTeams().ProvisionObjects(tenant, template.ParentHierarchy, null, parser, new Core.Framework.Provisioning.Model.Configuration.ApplyConfiguration());
                }
            }
        }
        
        #endregion

        [TestMethod]
        public void CanProvisionObjects()
        {
#if !ONPREMISES

            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("This test requires a user credentials, cannot be run using app-only for now");

            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                var template = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

                template.ParentHierarchy.Teams.TeamTemplates.Add(new TeamTemplate { JsonTemplate = _jsonTemplate });
                template.ParentHierarchy.Teams.Teams.Add(_team);

                Provision(template);

                // Wait for groups to be provisioned
                Thread.Sleep(5000);

                // Verify if Teams have been provisioned
                foreach (var teamName in _teamNames)
                {
                    var teams = GetTeamsByDisplayName(teamName);
                    Assert.IsTrue(teams.HasValues);
                }
            }
#else
            Assert.Inconclusive();
#endif
        }

        [TestMethod]
        public void CanUpdateObjects()
        {
#if !ONPREMISES
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("This test requires a user credentials, cannot be run using app-only for now");

            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                // Prepare the hierarchy for provisioning
                var template = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

                template.ParentHierarchy.Teams.Teams.Add(_team);

                // Initial provisioning of a Team
                Provision(template);

                // Get the just provisioned Team
                var provisionedTeam = GetTeamsByDisplayName(_team.DisplayName)?.FirstOrDefault();
                if (provisionedTeam != null && provisionedTeam.HasValues)
                {
                    // Store locally the just created Team ID
                    var teamId = provisionedTeam["id"].ToString();

                    // Now update the Team and test delta handling
                    template.ParentHierarchy.Teams.Teams[0].FunSettings.AllowGiphy =
                        !template.ParentHierarchy.Teams.Teams[0].FunSettings.AllowGiphy;
                    template.ParentHierarchy.Teams.Teams[0].GuestSettings.AllowCreateUpdateChannels =
                        !template.ParentHierarchy.Teams.Teams[0].GuestSettings.AllowCreateUpdateChannels;
                    template.ParentHierarchy.Teams.Teams[0].MemberSettings.AllowDeleteChannels =
                        !template.ParentHierarchy.Teams.Teams[0].MemberSettings.AllowDeleteChannels;
                    template.ParentHierarchy.Teams.Teams[0].MessagingSettings.AllowUserEditMessages =
                        !template.ParentHierarchy.Teams.Teams[0].MessagingSettings.AllowUserEditMessages;
                    template.ParentHierarchy.Teams.Teams[0].Channels[0].Description += " - Updated";
                    template.ParentHierarchy.Teams.Teams[0].Channels[0].Tabs.Add(new TeamTab
                    {
                        DisplayName = "OneNote Tab 2",
                        TeamsAppId = "0d820ecd-def2-4297-adad-78056cde7c78"
                    });

                    template.ParentHierarchy.Teams.Teams[0].Security.AllowToAddGuests = true;

                    Provision(template);

                    var team = GetTeamById(teamId);
                    var existingChannels = GetTeamChannels(teamId);
                    var existingTabs = GetTeamChannelTabs(teamId, existingChannels[1]["id"].ToString());
                    var allowToAddGuests = GetAllowToAddGuests(teamId);

                    Assert.AreEqual(team["funSettings"]["allowGiphy"].Value<Boolean>(),
                        template.ParentHierarchy.Teams.Teams[0].FunSettings.AllowGiphy);
                    Assert.AreEqual(team["guestSettings"]["allowCreateUpdateChannels"].Value<Boolean>(),
                        template.ParentHierarchy.Teams.Teams[0].GuestSettings.AllowCreateUpdateChannels);
                    Assert.AreEqual(team["memberSettings"]["allowDeleteChannels"].Value<Boolean>(),
                        template.ParentHierarchy.Teams.Teams[0].MemberSettings.AllowDeleteChannels);
                    Assert.AreEqual(team["messagingSettings"]["allowUserEditMessages"].Value<Boolean>(),
                        template.ParentHierarchy.Teams.Teams[0].MessagingSettings.AllowUserEditMessages);
                    Assert.AreEqual(existingChannels[1]["description"],
                        template.ParentHierarchy.Teams.Teams[0].Channels[0].Description);
                    Assert.IsTrue(existingTabs.Any(t => t["displayName"].ToString() == template.ParentHierarchy.Teams.Teams[0].Channels[0].Tabs[1].DisplayName));
                    Assert.IsTrue(allowToAddGuests);
                }
                else
                {
                    // If the Team wasn't created ... just fail
                    Assert.IsTrue(false);
                }
            }
#else
            Assert.Inconclusive();
#endif
        }
    }
}
#endif