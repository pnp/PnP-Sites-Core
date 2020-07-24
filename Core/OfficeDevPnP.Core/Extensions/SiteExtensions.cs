using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class SiteExtensions
    {
#if !ONPREMISES
        /// <summary>
        /// Retrieves the current value for the Site Classification of a Site Collection
        /// </summary>
        /// <param name="site">The target site</param>
        /// <param name="classificationValue">The new value for the Site Classification</param>
        /// <param name="accessToken">The OAuth Access Token to consume Microsoft Graph, required only for GROUP#0 site collections</param>
        /// <returns>The classification for the site</returns>
        public static void SetSiteClassification(this Site site, String classificationValue, String accessToken = null)
        {
            // Determine the modern site template
            var baseTemplateValue = site.RootWeb.GetBaseTemplateId();
            switch (baseTemplateValue)
            {
                // It is a "modern" team site
                case "GROUP#0":
                   
                    if (String.IsNullOrEmpty(accessToken))
                    {
                        throw new ArgumentNullException("accessToken");
                    }

                    // Ensure the GroupId value
                    site.EnsureProperty(s => s.GroupId);

                    // Update the Classification of the Office 365 Group
                    // PATCH https://graph.microsoft.com/beta/groups/{groupId}
                    string updateGroupUrl = $"{GraphHttpClient.MicrosoftGraphBetaBaseUri}groups/{site.GroupId}";
                    var updateGroupResult = GraphHttpClient.MakePatchRequestForString(
                        updateGroupUrl,
                        content: new
                        {
                            classification = classificationValue
                        },
                        contentType: "application/json",
                        accessToken: accessToken);

                    // Still update the local value to give prompt feedback to the user
                    site.Classification = classificationValue;
                    site.Context.ExecuteQueryRetry();

                    break;
                // It is a "modern" communication site
                case "SITEPAGEPUBLISHING#0":
                default:

                    site.Classification = classificationValue;
                    site.Context.ExecuteQueryRetry();

                    break;
            }
        }

        /// <summary>
        /// Retrieves the current value for the Site Classification of a Site Collection
        /// </summary>
        /// <param name="site">The target site</param>
        /// <returns>The classification for the site</returns>
        public static string GetSiteClassification(this Site site)
        {
            site.EnsureProperty(s => s.Classification);
            return (site.Classification);
        }

        /// <summary>
        /// Checks if the current Site Collection is a "modern" Communication Site
        /// </summary>
        /// <param name="site">The target site</param>
        /// <returns>Returns true if the site is a Communication Site</returns>
        public static Boolean IsCommunicationSite(this Site site)
        {
            // First of all check if the site is full Communication Site
            var templateId = site.RootWeb.GetBaseTemplateId();

            var result = (templateId == "SITEPAGEPUBLISHING#0");

            if (!result)
            {
                // Otherwise check if the Communication Site feature is enabled
                var commSiteFeatureId = new Guid("f39dad74-ea79-46ef-9ef7-fe2370754f6f");
                result = site.RootWeb.IsFeatureActive(commSiteFeatureId);
            }

            return (result);
        }

        /// <summary>
        /// Checks if the current Site Collection is a "modern" Team Site
        /// </summary>
        /// <param name="site">The target site</param>
        /// <returns>Returns true if the site is a Team Site</returns>
        public static Boolean IsModernTeamSite(this Site site)
        {
            // First of all check if the site is full Team Site
            var templateId = site.RootWeb.GetBaseTemplateId();

            var result = (templateId == "GROUP#0");

            return (result);
        }

        /// <summary>
        /// Checks if this site collection has a Teams team linked
        /// </summary>
        /// <param name="site">Site collection</param>
        /// <param name="accessToken">Graph access token (groups.read.all) </param>
        /// <returns>True if there's a team</returns>
        public static bool HasTeamsTeam(this Site site, string accessToken)
        {
            bool result = false;

            site.EnsureProperties(s => s.RootWeb, s=>s.GroupId);

            // A site without a group cannot have been teamified
            if (site.GroupId == Guid.Empty)
            {
                return false;
            }

            // fall back to Graph untill we've a SharePoint approach that works
            result = UnifiedGroupsUtility.HasTeamsTeam(site.GroupId.ToString(), accessToken);

            // Problem is that this folder property is not always set
            /*
            site.EnsureProperties(s => s.RootWeb, s => s.GroupId);
            List defaultDocumentLibrary = site.RootWeb.DefaultDocumentLibrary();
            site.RootWeb.Context.Load(defaultDocumentLibrary, f=>f.RootFolder);
            site.RootWeb.Context.ExecuteQueryRetry();

            if (defaultDocumentLibrary.RootFolder.FolderExists("General"))
            {
                // Load folder properties
                var generalFolder = defaultDocumentLibrary.RootFolder.EnsureFolder("General", p => p.Properties);
                site.RootWeb.Context.Load(generalFolder);
                site.RootWeb.Context.ExecuteQueryRetry();

                // Do we have the Teams channel entry ?
                string Vti_TeamChannelUrl = "vti_teamchannelurl";

                if (generalFolder.Properties.FieldValues.ContainsKey(Vti_TeamChannelUrl))
                {
                    var teamChannelUrl = generalFolder.Properties.FieldValues[Vti_TeamChannelUrl]?.ToString();
                    if (!string.IsNullOrEmpty(teamChannelUrl))
                    {
                        // Sample teams url: https://teams.microsoft.com/l/channel/19%3A0000866a32964362b5db23f21f81704c%40thread.skype/General?groupId=c1430c5f-c423-44b8-b083-bd81ca3f09d0&tenantId=ad20b775-5d3b-40f5-b144-c5c2c772b73e
                        // Just verify the url has a reference to the site's group id
                        if (teamChannelUrl.ToLower().Contains(site.GroupId.ToString().ToLower()))
                        {
                            return true;
                        }
                    }
                }
            }
            */

            return result;
        }
        
#endif
    }
}
