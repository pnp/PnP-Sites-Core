using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = OfficeDevPnP.Core.Framework.Provisioning.Model.File;
using Microsoft.SharePoint.Client.UserProfiles;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{
#if !SP2013
    internal static class LocalizationExtensions
    {
        internal static void LocalizeWebParts(this Page page, Web web, TokenParser parser, PnPMonitoredScope scope)
        {
            var url = page.Url;
            var webParts = page.WebParts;
            LocalizeParts(web, parser, url, webParts, scope);
        }

        internal static void LocalizeWebParts(this File file, Web web, TokenParser parser, Microsoft.SharePoint.Client.File targetFile, PnPMonitoredScope scope)
        {
            var url = targetFile.ServerRelativeUrl;
            var webParts = file.WebParts;
            LocalizeParts(web, parser, url, webParts, scope);
        }

        /// <summary>
        /// Checks if replacing the "Accept-Language" header is allowed. This approach does not work when the user that's
        /// making the call has set one or more languages in their user profile
        /// </summary>
        /// <param name="web">Current site</param>
        /// <returns>True if the "Accept-Language" header approach can be used</returns>
        private static bool CanUseAcceptLanguageHeaderForLocalization(Web web)
        {
            if (web.Context.IsAppOnly())
            {
                return true;
            }

            var currentUser = web.EnsureProperty(w => w.CurrentUser);
            PeopleManager peopleManager = new PeopleManager(web.Context);
            var languageSettings = peopleManager.GetUserProfilePropertyFor(web.CurrentUser.LoginName, "SPS-MUILanguages");
            web.Context.ExecuteQueryRetry();

            if (languageSettings == null || String.IsNullOrEmpty(languageSettings.Value))
            {
                return true;
            }

            return false;
        }

        private static void LocalizeParts(Web web, TokenParser parser, string url, WebPartCollection webParts, PnPMonitoredScope scope)
        {
            if (CanUseAcceptLanguageHeaderForLocalization(web))
            {
                var context = web.Context;
                var allParts = web.GetWebParts(parser.ParseString(url)).ToList();
                foreach (var webPart in webParts)
                {
#if !SP2016
                    var partOnPage = allParts.FirstOrDefault(w => w.ZoneId == webPart.Zone && w.WebPart.ZoneIndex == webPart.Order);
#else
                    var partOnPage = allParts.FirstOrDefault(w => w.WebPart.ZoneIndex == webPart.Order);
#endif
                    if (webPart.Title.ContainsResourceToken() && partOnPage != null)
                    {
                        var resourceValues = parser.GetResourceTokenResourceValues(webPart.Title);
                        foreach (var resourceValue in resourceValues)
                        {
                            // Save property with correct locale on the request to make it stick
                            // http://sadomovalex.blogspot.no/2015/09/localize-web-part-titles-via-client.html
                            context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                            partOnPage.WebPart.Properties["Title"] = resourceValue.Item2;
                            partOnPage.SaveWebPartChanges();
                            context.ExecuteQueryRetry();
                        }
                    }
                }
                context.PendingRequest.RequestExecutor.WebRequest.Headers.Remove("Accept-Language");
            }
            else
            {
                // warning
                scope.LogWarning(CoreResources.Provisioning_Extensions_WebPartLocalization_Skip);
            }
        }

        internal static void LocalizeView(this Microsoft.SharePoint.Client.View view, Web web, string token, TokenParser parser, PnPMonitoredScope scope)
        {
            if (CanUseAcceptLanguageHeaderForLocalization(web))
            {
                var context = web.Context;
                var resourceValues = parser.GetResourceTokenResourceValues(token);
                foreach (var resourceValue in resourceValues)
                {
                    // Save property with correct locale on the request to make it stick
                    // http://sadomovalex.blogspot.no/2015/09/localize-web-part-titles-via-client.html
                    context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                    view.Title = resourceValue.Item2;
                    view.Update();
                    context.ExecuteQueryRetry();
                }
            }
            else
            {
                // warning
                scope.LogWarning(CoreResources.Provisioning_Extensions_ViewLocalization_Skip);
            }
        }

        internal static void LocalizeNavigationNode(this Microsoft.SharePoint.Client.NavigationNode navigationNode, Web web, string token, TokenParser parser, PnPMonitoredScope scope)
        {
            if (CanUseAcceptLanguageHeaderForLocalization(web))
            {
                var context = web.Context;
                var resourceValues = parser.GetResourceTokenResourceValues(token);
                foreach (var resourceValue in resourceValues)
                {
                    // Save property with correct locale on the request to make it stick
                    // http://sadomovalex.blogspot.no/2015/09/localize-web-part-titles-via-client.html
                    context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                    navigationNode.Title = resourceValue.Item2;
                    navigationNode.Update();
                    context.ExecuteQueryRetry();
                }
            }
            else
            {
                // warning
                scope.LogWarning(CoreResources.Provisioning_Extensions_ViewLocalization_Skip);
            }
        }
    }
#endif
}
