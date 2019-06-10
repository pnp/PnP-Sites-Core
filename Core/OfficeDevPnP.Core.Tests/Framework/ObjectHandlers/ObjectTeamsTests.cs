#if !ONPREMISES
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
        private readonly List<string> _teamNames = new List<string>();
        private string _jsonTemplate;
        private Team _team;
        private string _existingTeamId;

        [TestInitialize]
        public void Initialize()
        {
            const string teamTemplateName = "Sample Engineering Team";
            _teamNames.Add(teamTemplateName);
            _jsonTemplate = "{ \"template@odata.bind\": \"https://graph.microsoft.com/beta/teamsTemplates(\'standard\')\", \"visibility\": \"Private\", \"displayName\": \"" + teamTemplateName + "\", \"description\": \"This is a sample engineering team, used to showcase the range of properties supported by this API\", \"channels\": [ { \"displayName\": \"Announcements 📢\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements.\" }, { \"displayName\": \"Training 🏋️\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample training channel, that is favorited by default, and contains an example of pinned website and YouTube tabs.\", \"tabs\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.web\')\", \"name\": \"A Pinned Website\", \"configuration\": { \"contentUrl\": \"https://docs.microsoft.com/en-us/microsoftteams/microsoft-teams\" } }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.youtube\')\", \"name\": \"A Pinned YouTube Video\", \"configuration\": { \"contentUrl\": \"https://tabs.teams.microsoft.com/Youtube/Home/YoutubeTab?videoId=X8krAMdGvCQ\", \"websiteUrl\": \"https://www.youtube.com/watch?v=X8krAMdGvCQ\" } } ] }, { \"displayName\": \"Planning 📅 \", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\", \"isFavoriteByDefault\": false }, { \"displayName\": \"Issues and Feedback 🐞\", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\" } ], \"memberSettings\": { \"allowCreateUpdateChannels\": true, \"allowDeleteChannels\": true, \"allowAddRemoveApps\": true, \"allowCreateUpdateRemoveTabs\": true, \"allowCreateUpdateRemoveConnectors\": true }, \"guestSettings\": { \"allowCreateUpdateChannels\": false, \"allowDeleteChannels\": false }, \"funSettings\": { \"allowGiphy\": true, \"giphyContentRating\": \"Moderate\", \"allowStickersAndMemes\": true, \"allowCustomMemes\": true }, \"messagingSettings\": { \"allowUserEditMessages\": true, \"allowUserDeleteMessages\": true, \"allowOwnerDeleteMessages\": true, \"allowTeamMentions\": true, \"allowChannelMentions\": true }, \"installedApps\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.vsts\')\" }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'1542629c-01b3-4a6d-8f76-1938b779e48d\')\" } ] }";

            const string teamName = @"Unallowed &_,.;:/\""!@$%^*[]+=&'|<>{}-()?#¤`´~¨ Ääkköset";
            _teamNames.Add(teamName);
            var security = new TeamSecurity
            {
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
            
            // For testing updating
            _existingTeamId = ConfigurationManager.AppSettings["ExistingTeamId"];
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                foreach (var teamName in _teamNames)
                {
                    var teams = GetTeamsByDisplayName(teamName);

                    foreach (var team in teams)
                    {
                        DeleteTeam(team["id"].ToString());
                    }
                }
            }
        }

        private static JToken GetTeamsByDisplayName(string displayName)
        {
            if (displayName == null) return null;

            var accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.Read.All");

            var requestUrl = $"https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '{HttpUtility.UrlEncode(displayName.Replace("'", "''"))}'";

            var response = HttpHelper.MakeGetRequestForString(requestUrl, accessToken);
            var json = JToken.Parse(response);

            return json["value"];
        }

        private static void DeleteTeam(string id)
        {
            var accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.ReadWrite.All");

            var requestUrl = $"https://graph.microsoft.com/v1.0/groups/{id}";

            HttpHelper.MakeDeleteRequest(requestUrl, accessToken);
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
#if !ONPREMISES
            var template = new ProvisioningTemplate {ParentHierarchy = new ProvisioningHierarchy()};

            template.ParentHierarchy.Teams.TeamTemplates.Add(new TeamTemplate { JsonTemplate = _jsonTemplate });
            template.ParentHierarchy.Teams.Teams.Add(_team);

            Provision(template);

            Assert.IsTrue(TeamsHaveBeenProvisioned());
#else
            Assert.Inconclusive();
#endif
        }

        [TestMethod]
        public void CanUpdateObjects()
        {
            var template = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            template.ParentHierarchy.Teams.Teams.Add(_team);

            template.ParentHierarchy.Teams.Teams[0].GroupId = _existingTeamId;

            Provision(template);

            Assert.IsTrue(TeamsHaveBeenUpdated());
        }

        private static void Provision(ProvisioningTemplate template)
        {
            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                using (var ctx = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(ctx);
                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectTeams().ProvisionObjects(tenant, template.ParentHierarchy, null, parser, new ProvisioningTemplateApplyingInformation());
                }
            }
        }

        private bool TeamsHaveBeenProvisioned()
        {
            // Wait for groups to be provisioned
            Thread.Sleep(5000);

            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                foreach (var teamName in _teamNames)
                {
                    var teams = GetTeamsByDisplayName(teamName);
                    if (!teams.HasValues) return false;
                }
            }

            return true;
        }

        private bool TeamsHaveBeenUpdated()
        {
            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                var accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.Read.All");

                var existingChannels = ObjectTeams.GetExistingTeamChannels(_existingTeamId, accessToken);

                var channels = _team.Channels;

                foreach (var channel in channels)
                {
                    var existingChannel = existingChannels.FirstOrDefault(x => x["displayName"].ToString() == channel.DisplayName);

                    if (existingChannel == null || channel.Description != existingChannel["description"].ToString()) return false;

                    var existingTabs = ObjectTeams.GetExistingTeamChannelTabs(_existingTeamId, existingChannel["id"].ToString(), accessToken);

                    foreach (var tab in channel.Tabs)
                    {
                        var existingTab = existingTabs.FirstOrDefault(x => HttpUtility.UrlDecode(x["displayName"].ToString()) == tab.DisplayName && x["teamsAppId"].ToString() == tab.TeamsAppId);

                        if (existingTab == null) return false;

                        // todo: check tab configurations after tab resource provisioning has been implemented and included in the test
                    }
                }
            }

            return true;
        }
    }
}
#endif