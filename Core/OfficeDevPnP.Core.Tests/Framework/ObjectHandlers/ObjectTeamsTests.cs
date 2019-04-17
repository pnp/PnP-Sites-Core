using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
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

        [TestInitialize]
        public void Initialize()
        {
            _teamNames.Add("Unit Test");
            _teamNames.Add("Sample Engineering Team");
            _teamTemplates.Add("{ \"template@odata.bind\": \"https://graph.microsoft.com/beta/teamsTemplates(\'standard\')\", \"displayName\": \"" + _teamNames[0] + "\", \"description\": \"Unit test\" }");
            _teamTemplates.Add("{ \"template@odata.bind\": \"https://graph.microsoft.com/beta/teamsTemplates(\'standard\')\", \"visibility\": \"Private\", \"displayName\": \"" + _teamNames[1] + "\", \"description\": \"This is a sample engineering team, used to showcase the range of properties supported by this API\", \"channels\": [ { \"displayName\": \"Announcements 📢\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements.\" }, { \"displayName\": \"Training 🏋️\", \"isFavoriteByDefault\": true, \"description\": \"This is a sample training channel, that is favorited by default, and contains an example of pinned website and YouTube tabs.\", \"tabs\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.web\')\", \"name\": \"A Pinned Website\", \"configuration\": { \"contentUrl\": \"https://docs.microsoft.com/en-us/microsoftteams/microsoft-teams\" } }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.youtube\')\", \"name\": \"A Pinned YouTube Video\", \"configuration\": { \"contentUrl\": \"https://tabs.teams.microsoft.com/Youtube/Home/YoutubeTab?videoId=X8krAMdGvCQ\", \"websiteUrl\": \"https://www.youtube.com/watch?v=X8krAMdGvCQ\" } } ] }, { \"displayName\": \"Planning 📅 \", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\", \"isFavoriteByDefault\": false }, { \"displayName\": \"Issues and Feedback 🐞\", \"description\": \"This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.\" } ], \"memberSettings\": { \"allowCreateUpdateChannels\": true, \"allowDeleteChannels\": true, \"allowAddRemoveApps\": true, \"allowCreateUpdateRemoveTabs\": true, \"allowCreateUpdateRemoveConnectors\": true }, \"guestSettings\": { \"allowCreateUpdateChannels\": false, \"allowDeleteChannels\": false }, \"funSettings\": { \"allowGiphy\": true, \"giphyContentRating\": \"Moderate\", \"allowStickersAndMemes\": true, \"allowCustomMemes\": true }, \"messagingSettings\": { \"allowUserEditMessages\": true, \"allowUserDeleteMessages\": true, \"allowOwnerDeleteMessages\": true, \"allowTeamMentions\": true, \"allowChannelMentions\": true }, \"installedApps\": [ { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'com.microsoft.teamspace.tab.vsts\')\" }, { \"teamsApp@odata.bind\": \"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps(\'1542629c-01b3-4a6d-8f76-1938b779e48d\')\" } ] }");
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(AcquireTokenAsync(resource, scope))))
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
            var accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.Read.All");

            var requestUrl = $"https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '{displayName}'";

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
            var template = new ProvisioningTemplate {ParentHierarchy = new ProvisioningHierarchy()};

            foreach (var teamTemplate in _teamTemplates)
            {
                template.ParentHierarchy.Teams.TeamTemplates.Add(new TeamTemplate { JsonTemplate = teamTemplate });
            }

            using (new PnPProvisioningContext((resource, scope) => Task.FromResult(AcquireTokenAsync(resource, scope))))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectTeams().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                }

                Assert.IsTrue(TeamsHaveBeenProvisioned());
            }
        }

        private bool TeamsHaveBeenProvisioned()
        {
            foreach (var teamName in _teamNames)
            {
                var teams = GetTeamsByDisplayName(teamName);
                if (!teams.HasValues) return false;
            }

            return true;
        }

        private static string AcquireTokenAsync(string resource, string scope = null)
        {
            var tenantId = GetTenantIdByUrl(TestCommon.AppSetting("SPOTenantUrl"));
            if (tenantId == null) return null;

            var clientId = TestCommon.AppSetting("AppId");
            var clientSecret = TestCommon.AppSetting("AppSecret");
            var username = TestCommon.AppSetting("SPOUserName");
            var password = TestCommon.AppSetting("SPOPassword");

            string body;
            string response;
            if (scope == null) // use v1 endpoint
            {
                body = $"grant_type=password&client_id={clientId}&username={username}&password={password}&resource={resource}&client_secret={WebUtility.UrlEncode(clientSecret)}";
                response = HttpHelper.MakePostRequestForString($"https://login.microsoftonline.com/{tenantId}/oauth2/token", body, "application/x-www-form-urlencoded");
            }
            else // use v2 endpoint
            {
                body = $"grant_type=password&client_id={clientId}&username={username}&password={password}&scope={scope}&client_secret={WebUtility.UrlEncode(clientSecret)}";
                response = HttpHelper.MakePostRequestForString($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", body, "application/x-www-form-urlencoded");
            }

            var json = JToken.Parse(response);
            return json["access_token"].ToString();
        }

        private static string GetTenantIdByUrl(string tenantUrl)
        {
            var tenantName = GetTenantNameFromUrl(tenantUrl);
            if (tenantName == null) return null;

            var url = $"https://login.microsoftonline.com/{tenantName}.onmicrosoft.com/.well-known/openid-configuration";
            var response = HttpHelper.MakeGetRequestForString(url);
            var json = JToken.Parse(response);

            var tokenEndpointUrl = json["token_endpoint"].ToString();
            return GetTenantIdFromAadEndpointUrl(tokenEndpointUrl);
        }

        private static string GetTenantNameFromUrl(string tenantUrl)
        {
            return GetSubstringFromMiddle(tenantUrl, "https://", "-admin.sharepoint.com");
        }

        private static string GetTenantIdFromAadEndpointUrl(string aadEndpointUrl)
        {
            return GetSubstringFromMiddle(aadEndpointUrl, "https://login.microsoftonline.com/", "/oauth2/");
        }

        private static string GetSubstringFromMiddle(string originalString, string prefix, string suffix)
        {
            var index = originalString.IndexOf(suffix, StringComparison.OrdinalIgnoreCase);
            return index != -1 ? originalString.Substring(prefix.Length, index - prefix.Length) : null;
        }
    }
}
