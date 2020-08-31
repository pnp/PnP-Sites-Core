using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Utilities.Async;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Themes
{
#if !ONPREMISES
    /// <summary>
    /// Extension class for the Web object useful to apply custom Themes
    /// </summary>
    public static class ThemeManager
    {
        /// <summary>
        /// Extension method to apply a Theme to a target web
        /// </summary>
        /// <param name="web"></param>
        /// <param name="jsonTheme"></param>
        /// <param name="themeName"></param>
        /// <returns></returns>
        public static Boolean ApplyTheme(this Web web, String jsonTheme, String themeName = null)
        {
            return Task.Run(() => ApplyThemeAsync(web, jsonTheme, themeName)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Extension method to apply OOTB SharePoint's theme to a target web
        /// </summary>
        /// <param name="web"></param>
        /// <param name="sharePointTheme"></param>
        /// <param name="themeName"></param>
        /// <returns></returns>
        public static Boolean ApplyTheme(this Web web, SharePointTheme sharePointTheme, string themeName = null)
        {
            string themeJsonString = GetThemeJsonAsString(sharePointTheme);
            themeName = themeName ?? sharePointTheme.ToString();
            return Task.Run(() => ApplySiteThemeAsync(web, themeJsonString, themeName)).GetAwaiter().GetResult();
        }

        public static async Task<Boolean> ApplyThemeAsync(Web web, String jsonTheme, String themeName = null)
        {
            if (web == null)
            {
                throw new ArgumentException(nameof(web));
            }
            if (string.IsNullOrEmpty(jsonTheme))
            {
                throw new ArgumentException(nameof(jsonTheme));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest((ClientContext)web.Context, ThemeAction.ApplyTheme,
                new
                {
                    name = themeName ?? String.Empty,
                    themeJson = JsonConvert.SerializeObject(new
                    {
                        palette = JsonConvert.DeserializeObject(jsonTheme)
                    }),
                });
        }

        internal static async Task<Boolean> ApplySiteThemeAsync(Web web, String jsonTheme, String themeName = null)
        {
            if (web == null)
            {
                throw new ArgumentException(nameof(web));
            }
            if (string.IsNullOrEmpty(jsonTheme))
            {
                throw new ArgumentException(nameof(jsonTheme));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest((ClientContext)web.Context, ThemeAction.ApplyTheme, jsonTheme);
        }

        /// <summary>
        /// Internal private method to process an HTTP request toward the ThemeManager REST APIs
        /// </summary>
        /// <param name="context">The current ClientContext of CSOM</param>
        /// <param name="action">The action to perform</param>
        /// <param name="postObject">The body of the request</param>
        /// <param name="accessToken">An optional Access Token for OAuth authorization</param>
        /// <returns>A boolean declaring whether the operation was successful</returns>
        private static async Task<bool> BaseRequest(ClientContext context, ThemeAction action, Object postObject, String accessToken = null)
        {
            var result = false;

            // If we don't have the access token
            if (String.IsNullOrEmpty(accessToken))
            {
                // Try to get one from the current context
                accessToken = context.GetAccessToken();
            }

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // We're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    // Reference here: https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-rest-api
                    string requestUrl = $"{context.Web.Url}/_api/thememanager/{action.ToString()}";

                    // Always make a POST request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("ACCEPT", "application/json; odata.metadata=minimal");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigestAsync());
                    request.Headers.Add("ODATA-VERSION", "4.0");

                    if (postObject != null)
                    {
                        var jsonBody = JsonConvert.SerializeObject(postObject);
                        var requestBody = new StringContent(jsonBody);
                        MediaTypeHeaderValue sharePointJsonMediaType;
                        MediaTypeHeaderValue.TryParse("application/json;charset=utf-8", out sharePointJsonMediaType);
                        requestBody.Headers.ContentType = sharePointJsonMediaType;
                        request.Content = requestBody;
                    }

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            try
                            {
                                var responseJson = JObject.Parse(responseString);
                                result = true;
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return await Task.Run(() => result);
        }

        /// <summary>
        /// Internal private method to process an HTTP request toward the ThemeManager REST APIs
        /// </summary>
        /// <param name="context">The current ClientContext of CSOM</param>
        /// <param name="action">The action to perform</param>
        /// <param name="postObject">The body of the request</param>
        /// <param name="accessToken">An optional Access Token for OAuth authorization</param>
        /// <returns>A boolean declaring whether the operation was successful</returns>
        private static async Task<bool> BaseRequest(ClientContext context, ThemeAction action, string postObject, String accessToken = null)
        {
            var result = false;

            // If we don't have the access token
            if (String.IsNullOrEmpty(accessToken))
            {
                // Try to get one from the current context
                accessToken = context.GetAccessToken();
            }

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // We're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    // Reference here: https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-rest-api
                    string requestUrl = $"{context.Web.Url}/_api/thememanager/{action.ToString()}";

                    // Always make a POST request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("ACCEPT", "application/json; odata.metadata=minimal");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigestAsync());
                    request.Headers.Add("ODATA-VERSION", "4.0");

                    if (!string.IsNullOrEmpty(postObject))
                    {
                        var jsonBody = JObject.Parse(postObject);
                        var requestBody = new StringContent(postObject);
                        MediaTypeHeaderValue sharePointJsonMediaType;
                        MediaTypeHeaderValue.TryParse("application/json;charset=utf-8", out sharePointJsonMediaType);
                        requestBody.Headers.ContentType = sharePointJsonMediaType;
                        request.Content = requestBody;
                    }

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            try
                            {
                                var responseJson = JObject.Parse(responseString);
                                result = true;
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return await Task.Run(() => result);
        }

        internal static string GetThemeJsonAsString(SharePointTheme theme)
        {
            switch (theme)
            {
                case SharePointTheme.Blue:
                    return "{'name':'Blue','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":0,\"G\":120,\"B\":212,\"A\":255},\"themeLighterAlt\":{\"R\":239,\"G\":246,\"B\":252,\"A\":255},\"themeLighter\":{\"R\":222,\"G\":236,\"B\":249,\"A\":255},\"themeLight\":{\"R\":199,\"G\":224,\"B\":244,\"A\":255},\"themeTertiary\":{\"R\":113,\"G\":175,\"B\":229,\"A\":255},\"themeSecondary\":{\"R\":43,\"G\":136,\"B\":216,\"A\":255},\"themeDarkAlt\":{\"R\":16,\"G\":110,\"B\":190,\"A\":255},\"themeDark\":{\"R\":0,\"G\":90,\"B\":158,\"A\":255},\"themeDarker\":{\"R\":0,\"G\":69,\"B\":120,\"A\":255},\"accent\":{\"R\":135,\"G\":100,\"B\":184,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
                case SharePointTheme.Orange:
                    return "{'name':'Orange','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":202,\"G\":80,\"B\":16,\"A\":255},\"themeLighterAlt\":{\"R\":253,\"G\":247,\"B\":244,\"A\":255},\"themeLighter\":{\"R\":246,\"G\":223,\"B\":210,\"A\":255},\"themeLight\":{\"R\":239,\"G\":196,\"B\":173,\"A\":255},\"themeTertiary\":{\"R\":223,\"G\":143,\"B\":100,\"A\":255},\"themeSecondary\":{\"R\":208,\"G\":98,\"B\":40,\"A\":255},\"themeDarkAlt\":{\"R\":181,\"G\":73,\"B\":15,\"A\":255},\"themeDark\":{\"R\":153,\"G\":62,\"B\":12,\"A\":255},\"themeDarker\":{\"R\":113,\"G\":45,\"B\":9,\"A\":255},\"accent\":{\"R\":152,\"G\":111,\"B\":11,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
                case SharePointTheme.Red:
                    return "{'name':'Red','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":164,\"G\":38,\"B\":44,\"A\":255},\"themeLighterAlt\":{\"R\":251,\"G\":244,\"B\":244,\"A\":255},\"themeLighter\":{\"R\":240,\"G\":211,\"B\":212,\"A\":255},\"themeLight\":{\"R\":227,\"G\":175,\"B\":178,\"A\":255},\"themeTertiary\":{\"R\":200,\"G\":108,\"B\":112,\"A\":255},\"themeSecondary\":{\"R\":174,\"G\":56,\"B\":62,\"A\":255},\"themeDarkAlt\":{\"R\":147,\"G\":34,\"B\":39,\"A\":255},\"themeDark\":{\"R\":124,\"G\":29,\"B\":33,\"A\":255},\"themeDarker\":{\"R\":91,\"G\":21,\"B\":25,\"A\":255},\"accent\":{\"R\":202,\"G\":80,\"B\":16,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
                case SharePointTheme.Purple:
                    return "{'name':'Purple','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":135,\"G\":100,\"B\":184,\"A\":255},\"themeLighterAlt\":{\"R\":249,\"G\":248,\"B\":252,\"A\":255},\"themeLighter\":{\"R\":233,\"G\":226,\"B\":244,\"A\":255},\"themeLight\":{\"R\":215,\"G\":201,\"B\":234,\"A\":255},\"themeTertiary\":{\"R\":178,\"G\":154,\"B\":212,\"A\":255},\"themeSecondary\":{\"R\":147,\"G\":114,\"B\":192,\"A\":255},\"themeDarkAlt\":{\"R\":121,\"G\":89,\"B\":165,\"A\":255},\"themeDark\":{\"R\":102,\"G\":75,\"B\":140,\"A\":255},\"themeDarker\":{\"R\":75,\"G\":56,\"B\":103,\"A\":255},\"accent\":{\"R\":3,\"G\":131,\"B\":135,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
                case SharePointTheme.Green:
                    return "{'name':'Green','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":73,\"G\":130,\"B\":5,\"A\":255},\"themeLighterAlt\":{\"R\":246,\"G\":250,\"B\":240,\"A\":255},\"themeLighter\":{\"R\":219,\"G\":235,\"B\":199,\"A\":255},\"themeLight\":{\"R\":189,\"G\":218,\"B\":155,\"A\":255},\"themeTertiary\":{\"R\":133,\"G\":180,\"B\":76,\"A\":255},\"themeSecondary\":{\"R\":90,\"G\":145,\"B\":23,\"A\":255},\"themeDarkAlt\":{\"R\":66,\"G\":117,\"B\":5,\"A\":255},\"themeDark\":{\"R\":56,\"G\":99,\"B\":4,\"A\":255},\"themeDarker\":{\"R\":41,\"G\":73,\"B\":3,\"A\":255},\"accent\":{\"R\":3,\"G\":131,\"B\":135,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
                case SharePointTheme.Gray:
                    return "{'name':'Gray','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":105,\"G\":121,\"B\":126,\"A\":255},\"themeLighterAlt\":{\"R\":248,\"G\":249,\"B\":250,\"A\":255},\"themeLighter\":{\"R\":228,\"G\":233,\"B\":234,\"A\":255},\"themeLight\":{\"R\":205,\"G\":213,\"B\":216,\"A\":255},\"themeTertiary\":{\"R\":159,\"G\":173,\"B\":177,\"A\":255},\"themeSecondary\":{\"R\":120,\"G\":136,\"B\":141,\"A\":255},\"themeDarkAlt\":{\"R\":93,\"G\":108,\"B\":112,\"A\":255},\"themeDark\":{\"R\":79,\"G\":91,\"B\":95,\"A\":255},\"themeDarker\":{\"R\":58,\"G\":67,\"B\":70,\"A\":255},\"accent\":{\"R\":0,\"G\":120,\"B\":212,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
                case SharePointTheme.DarkYellow:
                    return "{'name':'Dark Yellow','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":255,\"G\":200,\"B\":61,\"A\":255},\"themeLighterAlt\":{\"R\":10,\"G\":8,\"B\":2,\"A\":255},\"themeLighter\":{\"R\":41,\"G\":32,\"B\":10,\"A\":255},\"themeLight\":{\"R\":77,\"G\":60,\"B\":18,\"A\":255},\"themeTertiary\":{\"R\":153,\"G\":120,\"B\":37,\"A\":255},\"themeSecondary\":{\"R\":224,\"G\":176,\"B\":54,\"A\":255},\"themeDarkAlt\":{\"R\":255,\"G\":206,\"B\":81,\"A\":255},\"themeDark\":{\"R\":255,\"G\":213,\"B\":108,\"A\":255},\"themeDarker\":{\"R\":255,\"G\":224,\"B\":146,\"A\":255},\"accent\":{\"R\":255,\"G\":200,\"B\":61,\"A\":255},\"neutralLighterAlt\":{\"R\":40,\"G\":40,\"B\":40,\"A\":255},\"neutralLighter\":{\"R\":49,\"G\":49,\"B\":49,\"A\":255},\"neutralLight\":{\"R\":63,\"G\":63,\"B\":63,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":72,\"G\":72,\"B\":72,\"A\":255},\"neutralQuaternary\":{\"R\":79,\"G\":79,\"B\":79,\"A\":255},\"neutralTertiaryAlt\":{\"R\":109,\"G\":109,\"B\":109,\"A\":255},\"neutralTertiary\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralSecondary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralPrimaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralPrimary\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"neutralDark\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"black\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"white\":{\"R\":31,\"G\":31,\"B\":31,\"A\":255},\"primaryBackground\":{\"R\":31,\"G\":31,\"B\":31,\"A\":255},\"primaryText\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"isInverted\":true,\"version\":\"\"}'}";
                case SharePointTheme.DarkBlue:
                    return "{'name':'Dark Blue','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":58,\"G\":150,\"B\":221,\"A\":255},\"themeLighterAlt\":{\"R\":2,\"G\":6,\"B\":9,\"A\":255},\"themeLighter\":{\"R\":9,\"G\":24,\"B\":35,\"A\":255},\"themeLight\":{\"R\":17,\"G\":45,\"B\":67,\"A\":255},\"themeTertiary\":{\"R\":35,\"G\":90,\"B\":133,\"A\":255},\"themeSecondary\":{\"R\":51,\"G\":133,\"B\":195,\"A\":255},\"themeDarkAlt\":{\"R\":75,\"G\":160,\"B\":225,\"A\":255},\"themeDark\":{\"R\":101,\"G\":174,\"B\":230,\"A\":255},\"themeDarker\":{\"R\":138,\"G\":194,\"B\":236,\"A\":255},\"accent\":{\"R\":58,\"G\":150,\"B\":221,\"A\":255},\"neutralLighterAlt\":{\"R\":29,\"G\":43,\"B\":60,\"A\":255},\"neutralLighter\":{\"R\":34,\"G\":50,\"B\":68,\"A\":255},\"neutralLight\":{\"R\":43,\"G\":61,\"B\":81,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":50,\"G\":68,\"B\":89,\"A\":255},\"neutralQuaternary\":{\"R\":55,\"G\":74,\"B\":95,\"A\":255},\"neutralTertiaryAlt\":{\"R\":79,\"G\":99,\"B\":122,\"A\":255},\"neutralTertiary\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralSecondary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralPrimaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralPrimary\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"neutralDark\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"black\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"white\":{\"R\":24,\"G\":37,\"B\":52,\"A\":255},\"primaryBackground\":{\"R\":24,\"G\":37,\"B\":52,\"A\":255},\"primaryText\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"isInverted\":true,\"version\":\"\"}'}";
                case SharePointTheme.Teal:
                    return "{'name':'Teal','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":3,\"G\":120,\"B\":124,\"A\":255},\"themeLighterAlt\":{\"R\":240,\"G\":249,\"B\":250,\"A\":255},\"themeLighter\":{\"R\":197,\"G\":233,\"B\":234,\"A\":255},\"themeLight\":{\"R\":152,\"G\":214,\"B\":216,\"A\":255},\"themeTertiary\":{\"R\":73,\"G\":174,\"B\":177,\"A\":255},\"themeSecondary\":{\"R\":19,\"G\":137,\"B\":141,\"A\":255},\"themeDarkAlt\":{\"R\":2,\"G\":109,\"B\":112,\"A\":255},\"themeDark\":{\"R\":2,\"G\":92,\"B\":95,\"A\":255},\"themeDarker\":{\"R\":1,\"G\":68,\"B\":70,\"A\":255},\"accent\":{\"R\":79,\"G\":107,\"B\":237,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
                default:
                    return "{'name':'Blue','themeJson':'{\"backgroundImageUri\":\"\",\"palette\":{\"themePrimary\":{\"R\":0,\"G\":120,\"B\":212,\"A\":255},\"themeLighterAlt\":{\"R\":239,\"G\":246,\"B\":252,\"A\":255},\"themeLighter\":{\"R\":222,\"G\":236,\"B\":249,\"A\":255},\"themeLight\":{\"R\":199,\"G\":224,\"B\":244,\"A\":255},\"themeTertiary\":{\"R\":113,\"G\":175,\"B\":229,\"A\":255},\"themeSecondary\":{\"R\":43,\"G\":136,\"B\":216,\"A\":255},\"themeDarkAlt\":{\"R\":16,\"G\":110,\"B\":190,\"A\":255},\"themeDark\":{\"R\":0,\"G\":90,\"B\":158,\"A\":255},\"themeDarker\":{\"R\":0,\"G\":69,\"B\":120,\"A\":255},\"accent\":{\"R\":135,\"G\":100,\"B\":184,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}},\"cacheToken\":\"\",\"isDefault\":true,\"version\":\"\"}'}";
            }
        }
    }

    /// <summary>
    /// Available actions for ThemeManager REST APIs
    /// </summary>
    public enum ThemeAction
    {
        /// <summary>
        /// Add a new Tenant Theme
        /// </summary>
        AddTenantTheme,
        /// <summary>
        /// Delete an existing Tenant Theme
        /// </summary>
        DeleteTenantTheme,
        /// <summary>
        /// Get the whole list of Tenant Themes
        /// </summary>
        GetTenantThemingOptions,
        /// <summary>
        /// Apply a Theme to the target web
        /// </summary>
        ApplyTheme,
        /// <summary>
        /// Update a Tenant Theme
        /// </summary>
        UpdateTenantTheme,
    }
#endif 
}
