using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// Static class full of helper methods to make HTTP requests
    /// </summary>
    public static class HttpHelper
    {
        public const string JsonContentType = "application/json";

        /// <summary>
        /// Static readonly instance of HttpClient to improve performances
        /// </summary>
        /// <remarks>
        /// See https://docs.microsoft.com/en-us/azure/architecture/antipatterns/improper-instantiation/
        /// </remarks>
        private static readonly HttpClient httpClient =
            new HttpClient(new HttpClientHandler { AllowAutoRedirect = true }, true);

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <returns>The String value of the result</returns>
        public static string MakeGetRequestForString(string requestUrl,
            string accessToken = null)
        {
            return (MakeHttpRequest<String>("GET",
                requestUrl,
                accessToken,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The accept header for the response</param>
        /// <param name="referer">The URL Referer for the request</param>
        /// <returns>The Stream  of the result</returns>
        public static System.IO.Stream MakeGetRequestForStream(string requestUrl,
            string accept,
            string accessToken = null,
            string referer = null)
        {
            return (MakeHttpRequest<System.IO.Stream>("GET",
                requestUrl,
                accessToken,
                accept: accept,
                referer: referer,
                resultPredicate: r => r.Content.ReadAsStreamAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="responseHeaders">The response headers of the HTTP request (output argument)</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The accept header for the response</param>
        /// <returns>The Stream  of the result</returns>
        public static System.IO.Stream MakeGetRequestForStreamWithResponseHeaders(string requestUrl,
            string accept,
            out HttpResponseHeaders responseHeaders,
            string accessToken = null)
        {
            return (MakeHttpRequest<System.IO.Stream>("GET",
                requestUrl,
                out responseHeaders,
                accessToken,
                accept: accept,
                resultPredicate: r => r.Content.ReadAsStreamAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP POST request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        public static void MakePostRequest(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null)
        {
            MakeHttpRequest<string>("POST",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType);
        }

        /// <summary>
        /// This helper method makes an HTTP POST request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <returns>The String value of the result</returns>
        public static string MakePostRequestForString(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null)
        {
            return (MakeHttpRequest<String>("POST",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        public static HttpResponseHeaders MakePostRequestForHeaders(string requestUrl, object content = null, string contentType = null, string accessToken = null)
        {
            return MakeHttpRequest("POST",
                requestUrl,
                accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: response => response.Headers);
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        public static void MakePutRequest(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null)
        {
            MakeHttpRequest<string>("PUT",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType);
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <returns>The String value of the result</returns>
        public static string MakePutRequestForString(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null)
        {
            return (MakeHttpRequest<String>("PUT",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP PATCH request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <returns>The String value of the result</returns>
        public static string MakePatchRequestForString(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null)
        {
            return (MakeHttpRequest<String>("PATCH",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP DELETE request
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <returns>The String value of the result</returns>
        public static void MakeDeleteRequest(string requestUrl,
            string accessToken = null)
        {
            MakeHttpRequest<string>("DELETE", requestUrl, accessToken);
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="referer">The URL Referer for the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        private static TResult MakeHttpRequest<TResult>(
            string httpMethod,
            string requestUrl,
            string accessToken = null,
            string accept = null,
            object content = null,
            string contentType = null,
            string referer = null,
            Func<HttpResponseMessage, TResult> resultPredicate = null)
        {
            HttpResponseHeaders responseHeaders;
            return (MakeHttpRequest<TResult>(httpMethod,
                requestUrl,
                out responseHeaders,
                accessToken,
                accept,
                content,
                contentType,
                referer,
                resultPredicate));
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="responseHeaders">The response headers of the HTTP request (output argument)</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="referer">The URL Referer for the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        private static TResult MakeHttpRequest<TResult>(
            string httpMethod,
            string requestUrl,
            out HttpResponseHeaders responseHeaders,
            string accessToken = null,
            string accept = null,
            object content = null,
            string contentType = null,
            string referer = null,
            Func<HttpResponseMessage, TResult> resultPredicate = null)
        {
            // Prepare the variable to hold the result, if any
            TResult result = default(TResult);
            responseHeaders = null;

            Uri requestUri = new Uri(requestUrl);

            // If we have the token, then handle the HTTP request

            // Set the Authorization Bearer token
            if (!string.IsNullOrEmpty(accessToken))
            {
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);
            }

            if (!string.IsNullOrEmpty(referer))
            {
                httpClient.DefaultRequestHeaders.Referrer = new Uri(referer);
            }

            // If there is an accept argument, set the corresponding HTTP header
            if (!string.IsNullOrEmpty(accept))
            {
                httpClient.DefaultRequestHeaders.Accept.Clear();
                httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue(accept));
            }

            // Prepare the content of the request, if any
            HttpContent requestContent = null;
            System.IO.Stream streamContent = content as System.IO.Stream;
            if (streamContent != null)
            {
                requestContent = new StreamContent(streamContent);
                requestContent.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            }
            else if (content != null)
            {
                var jsonString = content is string
                    ? content.ToString()
                    : JsonConvert.SerializeObject(content, Formatting.None, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        ContractResolver = new ODataBindJsonResolver(),

                    });
                requestContent = new StringContent(jsonString, Encoding.UTF8, contentType);
            }

            // Prepare the HTTP request message with the proper HTTP method
            HttpRequestMessage request = new HttpRequestMessage(
                new HttpMethod(httpMethod), requestUrl);

            // Set the request content, if any
            if (requestContent != null)
            {
                request.Content = requestContent;
            }

            // Fire the HTTP request
            HttpResponseMessage response = httpClient.SendAsync(request).Result;

            if (response.IsSuccessStatusCode)
            {
                // If the response is Success and there is a
                // predicate to retrieve the result, invoke it
                if (resultPredicate != null)
                {
                    result = resultPredicate(response);
                }

                // Get any response header and put it in the answer
                responseHeaders = response.Headers;
            }
            else
            {
                throw new ApplicationException(
                    string.Format("Exception while invoking endpoint {0}.", requestUrl),
                    new HttpException(
                        (int)response.StatusCode,
                        response.Content.ReadAsStringAsync().Result));
            }

            return (result);
        }
    }
}
