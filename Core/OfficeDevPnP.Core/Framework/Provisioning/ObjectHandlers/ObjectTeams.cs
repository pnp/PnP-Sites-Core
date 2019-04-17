using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Linq;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Object Handler to manage Microsoft Teams stuff
    /// </summary>
    internal class ObjectTeams : ObjectHandlerBase
    {
        public override string Name => "Teams";
        public override string InternalName => "Teams";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#if !ONPREMISES
            using (var scope = new PnPMonitoredScope(Name))
            {
                var accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.ReadWrite.All");

                // - Teams Templates
                var teamTemplates = template.ParentHierarchy.Teams?.TeamTemplates;
                if (teamTemplates != null && teamTemplates.Any())
                {
                    foreach (var teamTemplate in teamTemplates)
                    {
                        var team = CreateByTeamTemplate(scope, parser, teamTemplate, accessToken);
                        // possible further processing...
                    }
                }

                // - Teams
                // - Apps
            }
#endif

            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            // So far, no extraction
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#if !ONPREMISES
            if (!_willProvision.HasValue)
            {
                _willProvision = template.ParentHierarchy.Teams?.TeamTemplates?.Any() |
                    template.ParentHierarchy.Teams?.Teams?.Any();
            }
#else
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
#endif            
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }

        private static JToken CreateByTeamTemplate(PnPMonitoredScope scope, TokenParser parser, TeamTemplate teamTemplate, string accessToken)
        {
            HttpResponseHeaders responseHeaders;
            try
            {
                var content = parser.ParseString(teamTemplate.JsonTemplate);
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
    }
}
