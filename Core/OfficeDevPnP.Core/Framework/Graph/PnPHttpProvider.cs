using Microsoft.Graph;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using static Microsoft.SharePoint.Client.ClientContextExtensions;

namespace OfficeDevPnP.Core.Framework.Graph
{
    ///<summary>
    /// Class that deals with PnPHttpProvider methods
    ///</summary>  
    public class PnPHttpProvider : HttpProvider, IHttpProvider
    {
        private int _retryCount;
        private int _delay;
        private string _userAgent;

        /// <summary>
        /// Constructor for the PnPHttpProvider class
        /// </summary>
        /// <param name="retryCount">Maximum retry Count</param>
        /// <param name="delay">Delay Time</param>
        /// <param name="userAgent">User-Agent string to set</param>
        public PnPHttpProvider(int retryCount = 10, int delay = 500, string userAgent = null) : base()
        {
            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");

            if (delay <= 0)
                throw new ArgumentException("Provide a delay greater than zero.");

            this._retryCount = retryCount;
            this._delay = delay;
            this._userAgent = userAgent;
        }

        /// <summary>
        /// Custom implementation of the IHttpProvider.SendAsync method to handle retry logic
        /// </summary>
        /// <param name="request">The HTTP Request Message</param>
        /// <param name="completionOption">The completion option</param>
        /// <param name="cancellationToken">The cancellation token</param>
        /// <returns>The result of the asynchronous request</returns>
        /// <remarks>See here for further details: https://graph.microsoft.io/en-us/docs/overview/errors</remarks>
        Task<HttpResponseMessage> IHttpProvider.SendAsync(HttpRequestMessage request, HttpCompletionOption completionOption, CancellationToken cancellationToken)
        {
            // Retry logic variables
            int retryAttempts = 0;
            int backoffInterval = this._delay;

            HttpRequestMessage workrequest = request;

            // Loop until we need to retry
            while (retryAttempts < this._retryCount)
            {
                try
                {
                    // Add the PnP User Agent string
                    workrequest.Headers.UserAgent.TryParseAdd(string.IsNullOrEmpty(_userAgent) ? $"{PnPCoreUtilities.PnPCoreUserAgent}" : _userAgent);

                    // Make the request
                    Task<HttpResponseMessage> result = base.SendAsync(workrequest, completionOption, cancellationToken);

                    if (result != null && result.Result != null && (result.Result.StatusCode == (HttpStatusCode)429 || result.Result.StatusCode == (HttpStatusCode)503))
                    {
                        // And return the response in case of success
                        Log.Warning(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_SendAsyncRetry, $"{backoffInterval}");

                        //Add delay for retry
                        Task.Delay(backoffInterval).Wait();

                        //Add to retry count and increase delay.
                        retryAttempts++;
                        backoffInterval = backoffInterval * 2;

                        workrequest = workrequest.CloneRequest();
                    }
                    else
                    {
                        // And return the response in case of success
                        return result;
                    }
                }
                // Or handle any ServiceException
                catch (ServiceException ex)
                {
                    // Check if the is an InnerException
                    // And if it is a WebException
                    var wex = ex.InnerException as WebException;
                    if (wex != null)
                    {
                        var response = wex.Response as HttpWebResponse;
                        // Check if request was throttled - http status code 429
                        // Check is request failed due to server unavailable - http status code 503
                        if (response != null &&
                            (response.StatusCode == (HttpStatusCode) 429 || response.StatusCode == (HttpStatusCode) 503))
                        {
                            Log.Warning(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_SendAsyncRetry,
                                backoffInterval);

                            //Add delay for retry
                            Task.Delay(backoffInterval).Wait();

                            //Add to retry count and increase delay.
                            retryAttempts++;
                            backoffInterval = backoffInterval * 2;
                            workrequest = workrequest.CloneRequest();
                        }
                        else
                        {
                            Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_SendAsyncRetryException,
                                wex.ToString());
                            throw;
                        }
                    }
                    throw;
                }
            }

            throw new MaximumRetryAttemptedException($"Maximum retry attempts {this._retryCount}, has be attempted.");
        }
    }
}
