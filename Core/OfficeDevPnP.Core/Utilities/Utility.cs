using System;
using System.Linq.Expressions;
using Microsoft.SharePoint.Client;
using System.Net;

namespace OfficeDevPnP.Core.Utilities
{
    public static partial class Utility
    {

        public static string GetAccessTokenFromClientContext(ClientRuntimeContext clientContext)
        {
            string accessToken = null;
            // Issue a dummy request to get it from the Authorization header
            clientContext.ExecutingWebRequest += (s, e) =>
            {
                string authorization = e.WebRequestExecutor.RequestHeaders["Authorization"];
                accessToken = authorization.Replace("Bearer ", string.Empty);
            };
            clientContext.ExecuteQueryRetry();
            return accessToken;
        }

        /// <summary>
        /// Returns the healthscore for a SharePoint Server
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static int GetHealthScore(string url)
        {
            int value = 0;
            Uri baseUri = new Uri(url);
            Uri checkUri = new Uri(baseUri, "_layouts/15/blank.htm");
            WebRequest webRequest = WebRequest.Create(checkUri);
            webRequest.Method = "HEAD";
            webRequest.UseDefaultCredentials = true;
            using (WebResponse webResponse = webRequest.GetResponse())
            {
                foreach (string header in webResponse.Headers)
                {
                    if (header == "X-SharePointHealthScore")
                    {
                        value = Int32.Parse(webResponse.Headers[header]);
                        break;
                    }
                }
            }
            return value;
        }
    }
}
