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
        
#endif
    }
}
