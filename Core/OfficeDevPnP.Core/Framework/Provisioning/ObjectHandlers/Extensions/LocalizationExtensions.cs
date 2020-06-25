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
using System.Globalization;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{

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

#if SP2019
            // TODO: this does not work for HighTrust Auth with user delegation
            // https://docs.microsoft.com/en-us/previous-versions/office/mt210897(v=office.15)#app-only-policy-authorization
            // You can't use the app-only policy with the following APIs:
            // User Profile, ...

            /*
            var currentUser = web.EnsureProperty(w => w.CurrentUser);
            var currentUser = web.EnsureProperty(w => w.CurrentUser);
            PeopleManager peopleManager = new PeopleManager(web.Context);
            PeopleManager peopleManager = new PeopleManager(web.Context);
            
            var userProfilePropertiesForUser = new Microsoft.SharePoint.Client.UserProfiles.UserProfilePropertiesForUser(
                web.Context,
                web.CurrentUser.LoginName,
                new string[] {"SPS-MUILanguages"}
                );
            var propertyArray = peopleManager.GetUserProfilePropertiesFor(userProfilePropertiesForUser);
            web.Context.Load(userProfilePropertiesForUser);
            web.Context.ExecuteQueryRetry();
            string languageSettings = null;
            languageSettings = propertyArray.FirstOrDefault();
            if (string.IsNullOrEmpty(languageSettings))
            {
                return true;
            }
            */

            var testCurrentUser = web.EnsureProperty(w => w.CurrentUser);
            testCurrentUser.EnsureProperty(u => u.LoginName);

            // currentUser.LoginName

            /*
            PeopleManager testPeopleManager = new PeopleManager(web.Context);
            var testProperties = testPeopleManager.GetMyProperties();
            web.Context.ExecuteQueryRetry();
            */


            if (web.Context.IsAppOnlyWithDelegation())
            {
                return true;
            }
#endif
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
                web.EnsureProperties(w => w.Language, w => w.IsMultilingual, w => w.SupportedUILanguageIds);
                if (web.IsMultilingual)
                {
                    //just update if web is multilingual 
                    var allParts = web.GetWebParts(parser.ParseString(url)).ToList();
                    foreach (var webPart in webParts)
                    {
                        var partOnPage = allParts.FirstOrDefault(w => w.ZoneId == webPart.Zone && w.WebPart.ZoneIndex == webPart.Order);
                        if (webPart.Title.ContainsResourceToken() && partOnPage != null)
                        {
                            var resourceValues = parser.GetResourceTokenResourceValues(webPart.Title);
                            foreach (var resourceValue in resourceValues)
                            {
                                var translationculture = new CultureInfo(resourceValue.Item1);
                                if (web.SupportedUILanguageIds.Contains(translationculture.LCID))
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
                    }
                    context.PendingRequest.RequestExecutor.WebRequest.Headers.Remove("Accept-Language");
                }
                else
                {
                    //skip since web is not multilingual
                    scope.LogWarning(CoreResources.Provisioning_Extensions_WebPartLocalization_NoMUI_Skip);
                }
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
                //preserve Default Language to set last since otherwise Entries in QuickLaunch can show wrong language
                web.EnsureProperties(w => w.Language, w => w.IsMultilingual, w => w.SupportedUILanguageIds);
                if (web.IsMultilingual)
                {
                    //just update if web is multilingual 
                    var culture = new CultureInfo((int)web.Language);

                    var resourceValues = parser.GetResourceTokenResourceValues(token);

                    var defaultLanguageResource = resourceValues.FirstOrDefault(r => r.Item1.Equals(culture.Name, StringComparison.InvariantCultureIgnoreCase));
                    if (defaultLanguageResource != null)
                    {
                        foreach (var resourceValue in resourceValues.Where(r => !r.Item1.Equals(culture.Name, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            var translationculture = new CultureInfo(resourceValue.Item1);
                            if (web.SupportedUILanguageIds.Contains(translationculture.LCID))
                            {
                                // Save property with correct locale on the request to make it stick
                                // http://sadomovalex.blogspot.no/2015/09/localize-web-part-titles-via-client.html
                                context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                                view.Title = resourceValue.Item2;
                                view.Update();
                                context.ExecuteQueryRetry();
                            }
                        }
                        //Set for default Language of Web
                        context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = defaultLanguageResource.Item1;
                        view.Title = defaultLanguageResource.Item2;
                        view.Update();
                        context.ExecuteQueryRetry();
                    }
                    else
                    {
                        //skip since default language of web is not contained in resource file
                        scope.LogWarning(CoreResources.Provisioning_Extensions_ViewLocalization_DefaultLngMissing_Skip);
                    }
                }
                else
                {
                    //skip since web is not multilingual
                    scope.LogWarning(CoreResources.Provisioning_Extensions_ViewLocalization_NoMUI_Skip);
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
                web.EnsureProperties(w => w.Language, w => w.IsMultilingual, w => w.SupportedUILanguageIds);
                if (web.IsMultilingual)
                {
                    var context = web.Context;
                    var resourceValues = parser.GetResourceTokenResourceValues(token);
                    foreach (var resourceValue in resourceValues)
                    {
                        var translationculture = new CultureInfo(resourceValue.Item1);
                        if (web.SupportedUILanguageIds.Contains(translationculture.LCID))
                        {
                            // Save property with correct locale on the request to make it stick
                            // http://sadomovalex.blogspot.no/2015/09/localize-web-part-titles-via-client.html
                            context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                            navigationNode.Title = resourceValue.Item2;
                            navigationNode.Update();
                            context.ExecuteQueryRetry();
                        }
                    }
                }
                else
                {
                    scope.LogWarning(CoreResources.Provisioning_Extensions_NavigationLocalization_NoMUI_Skip);
                }
            }
            else
            {
                // warning
                scope.LogWarning(CoreResources.Provisioning_Extensions_NavigationLocalization_Skip);
            }
        }
    }
}
