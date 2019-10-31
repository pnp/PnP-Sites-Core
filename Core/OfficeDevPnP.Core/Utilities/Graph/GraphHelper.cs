using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Graph
{
    internal static class GraphHelper
    {
        public const String MicrosoftGraphBaseURI = "https://graph.microsoft.com/";

        /// <summary>
        /// Helper method to create or update an object through the Microsoft Graph
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="method">The HTTP method to use</param>
        /// <param name="uri">The URI for the Graph request</param>
        /// <param name="content">The content of the Graph request</param>
        /// <param name="contentType">The content type of the Graph request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <param name="alreadyExistsErrorMessage">The error message token that identifies an already existing item</param>
        /// <param name="warningMessage">The warning message to log when the target item already exists</param>
        /// <param name="matchingFieldName">The name of a field to match an already existing instance of the target item</param>
        /// <param name="matchingFieldValue">The value of a field to match an already existing instance of the target item</param>
        /// <param name="errorMessage">The error message to log when the create or update action fails</param>
        /// <param name="canPatch">Defines whether a Patch HTTP request can be executed to update an already existing target item</param>
        /// <returns>The ID of the create or updated target item</returns>
        public static String CreateOrUpdateGraphObject(
            PnPMonitoredScope scope,
            HttpMethodVerb method,
            String uri,
            Object content,
            String contentType,
            String accessToken,
            String alreadyExistsErrorMessage,
            String warningMessage,
            String matchingFieldName,
            String matchingFieldValue,
            String errorMessage,
            Boolean canPatch
            )
        {
            try
            {
                String itemId = null;
                String json = null;
                HttpResponseHeaders responseHeaders;

                // Try to create the Graph object
                switch (method)
                {
                    case HttpMethodVerb.POST:
                        json = HttpHelper.MakePostRequestForString(uri, content, contentType, accessToken);
                        if (!string.IsNullOrEmpty(json))
                        {
                            itemId = JToken.Parse(json).Value<String>("id");
                        }                        
                        break;
                    case HttpMethodVerb.PUT:
                        json = HttpHelper.MakePutRequestForString(uri, content, contentType, accessToken);
                        if (!string.IsNullOrEmpty(json))
                        {
                            itemId = JToken.Parse(json).Value<String>("id");
                        }                        
                        break;
                    case HttpMethodVerb.POST_WITH_RESPONSE_HEADERS:
                        responseHeaders = HttpHelper.MakePostRequestForHeaders(uri, content, contentType, accessToken);
                        itemId = responseHeaders.Location.ToString().Split('\'')[1];
                        break;
                }

                // Return the ID of the just created item
                return itemId;
            }
            catch (Exception ex)
            {
                // In case of exception, let's see if the target item already exists
                if (!String.IsNullOrEmpty(alreadyExistsErrorMessage) &&
                    !String.IsNullOrEmpty(matchingFieldName) &&
                    !String.IsNullOrEmpty(matchingFieldValue) &&
                    ex.InnerException.Message.Contains(alreadyExistsErrorMessage))
                {
                    try
                    {
                        if (!String.IsNullOrEmpty(warningMessage))
                        {
                            scope.LogWarning(warningMessage);
                        }

                        // If it's a POST we need to look for any existing item
                        String id = null;

                        // In case of PUT we already have the id
                        if (method == HttpMethodVerb.POST || method == HttpMethodVerb.POST_WITH_RESPONSE_HEADERS)
                        {
                            // Filter by field and value specified
                            id = ItemAlreadyExists(uri, matchingFieldName, matchingFieldValue, accessToken);
                            uri = $"{uri}/{id}";
                        }
                        else
                        {
                            id = matchingFieldValue;
                        }

                        // Patch the item, if supported
                        if (canPatch)
                        {
                            HttpHelper.MakePatchRequestForString(uri, content, contentType, accessToken);
                        }

                        return id;
                    }
                    catch (Exception exUpdate)
                    {
                        if (!String.IsNullOrEmpty(errorMessage))
                        {
                            scope.LogError(errorMessage, exUpdate.Message);
                        }
                        return null;
                    }
                }
                else
                {
                    return (null);
                }
            }
        }

        public static string ItemAlreadyExists(string uri, string matchingFieldName, string matchingFieldValue, string accessToken)
        {
            string id;
            String json = HttpHelper.MakeGetRequestForString($"{uri}?$select=id&$filter={matchingFieldName}%20eq%20'{WebUtility.UrlEncode(matchingFieldValue)}'", accessToken);
            // Get the id of existing item
            var ids = GetIdsFromList(json);
            id = ids.Length > 0 ? ids[0] : null;
            return id;
        }

        /// <summary>
        /// Retrieves the IDs of items in a JSON list
        /// </summary>
        /// <param name="json">The JSON list</param>
        /// <returns>The array of IDs</returns>
        public static string[] GetIdsFromList(string json)
        {
            return JsonConvert.DeserializeAnonymousType(json, new { value = new[] { new { id = "" } } }).value.Select(v => v.id).ToArray();
        }
    }

    internal enum HttpMethodVerb
    {
        GET,
        POST,
        PUT,
        PATCH,
        POST_WITH_RESPONSE_HEADERS
    }
}
