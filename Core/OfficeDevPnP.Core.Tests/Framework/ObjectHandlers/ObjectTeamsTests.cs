#if !ONPREMISES
using System.Collections.Generic;
using System.Configuration;
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
        private readonly List<string> _teamTemplates = new List<string>();
        private readonly List<Team> _teams = new List<Team>();

        [TestInitialize]
        public void Initialize()
        {
            _teamNames.Add("Unit Test");
            _teamTemplates.Add("{ \"template@odata.bind\": \"https://graph.microsoft.com/beta/teamsTemplates(\'standard\')\", \"displayName\": \"" + _teamNames[0] + "\", \"description\": \"Unit test\" }");

            _teamNames.Add("Sample Engineering Team");
            _teamTemplates.Add("{ \"template@odata.bind\": \"https://graph.microsoft.com/beta/teamsTemplates(\'standard\')\", \"visibility\": \"Private\", \"displayName\": \"" + _teamNames[1] + "\", \"description\": \"This is a sample engineering team, used to showcase the range of properties supported by this API\", \"channels\": [ { \"displayName\": \"Announcements 📢\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements.\" }, { \"displayName\": \"Training 🏋️\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample training channel, that is favorited by default, and contains an example of pinned website and YouTube tabs.\", \"tabs\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.web\')\", \"name\": \"A Pinned Website\", \"configuration\": { \"contentUrl\": \"https://docs.microsoft.com/en-us/microsoftteams/microsoft-teams\" } }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.youtube\')\", \"name\": \"A Pinned YouTube Video\", \"configuration\": { \"contentUrl\": \"https://tabs.teams.microsoft.com/Youtube/Home/YoutubeTab?videoId=X8krAMdGvCQ\", \"websiteUrl\": \"https://www.youtube.com/watch?v=X8krAMdGvCQ\" } } ] }, { \"displayName\": \"Planning 📅 \", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\", \"isFavoriteByDefault\": false }, { \"displayName\": \"Issues and Feedback 🐞\", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\" } ], \"memberSettings\": { \"allowCreateUpdateChannels\": true, \"allowDeleteChannels\": true, \"allowAddRemoveApps\": true, \"allowCreateUpdateRemoveTabs\": true, \"allowCreateUpdateRemoveConnectors\": true }, \"guestSettings\": { \"allowCreateUpdateChannels\": false, \"allowDeleteChannels\": false }, \"funSettings\": { \"allowGiphy\": true, \"giphyContentRating\": \"Moderate\", \"allowStickersAndMemes\": true, \"allowCustomMemes\": true }, \"messagingSettings\": { \"allowUserEditMessages\": true, \"allowUserDeleteMessages\": true, \"allowOwnerDeleteMessages\": true, \"allowTeamMentions\": true, \"allowChannelMentions\": true }, \"installedApps\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.vsts\')\" }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'1542629c-01b3-4a6d-8f76-1938b779e48d\')\" } ] }");

            _teamNames.Add(@"Unallowed &_,.;:/\""!@$%^*[]+=&'|<>{}-()?#¤`´~¨ Ääkköset");
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

            _teams.Add(new Team { DisplayName = _teamNames[2], Description = "Testing creating mailNickname from a display name that has unallowed and accented characters", Visibility = TeamVisibility.Public, Security = security, FunSettings = funSettings, GuestSettings = guestSettings, MemberSettings = memberSettings, MessagingSettings = messagingSettings});
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

            foreach (var teamTemplate in _teamTemplates)
            {
                template.ParentHierarchy.Teams.TeamTemplates.Add(new TeamTemplate { JsonTemplate = teamTemplate });
            }

            foreach (var team in _teams)
            {
                template.ParentHierarchy.Teams.Teams.Add(team);
            }

            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(TestCommon.AcquireTokenAsync(resource, scope))))
            {
                using (var ctx = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(ctx);
                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectTeams().ProvisionObjects(tenant, template.ParentHierarchy, null, parser, new ProvisioningTemplateApplyingInformation());
                }

                Assert.IsTrue(TeamsHaveBeenProvisioned());
            }
#else
            Assert.Inconclusive();
#endif
        }

        private bool TeamsHaveBeenProvisioned()
        {
            // Wait for groups to be provisioned
            Thread.Sleep(5000);

            foreach (var teamName in _teamNames)
            {
                var teams = GetTeamsByDisplayName(teamName);
                if (!teams.HasValues) return false;
            }

            return true;
        }
    }
}
#endif