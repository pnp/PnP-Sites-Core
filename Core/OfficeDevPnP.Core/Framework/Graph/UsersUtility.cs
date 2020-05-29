using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;
using OfficeDevPnP.Core.Diagnostics;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Graph
{
    /// <summary>
    /// Provides access to user operations in Microsoft Graph
    /// </summary>
    public static class UsersUtility
    {
        /// <summary>
        /// Returns the user with the provided userId from Azure Active Directory
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="userId">The unique identifier of the user in Azure Active Directory to return</param>        
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <returns>List with User objects</returns>
        public static Model.User GetUser(string accessToken, Guid userId, int startIndex = 0, int endIndex = 999, int retryCount = 10, int delay = 500)
        {
            return ListUsers(accessToken, $"id eq '{userId}'", null, startIndex, endIndex, retryCount, delay).FirstOrDefault();
        }

        /// <summary>
        /// Returns all the Users in the current domain
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>        
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <returns>List with User objects</returns>
        public static List<Model.User> ListUsers(string accessToken, int startIndex = 0, int endIndex = 999, int retryCount = 10, int delay = 500)
        {
            return ListUsers(accessToken, null, null, startIndex, endIndex, retryCount, delay);
        }

        /// <summary>
        /// Returns all the Users in the current domain filtered out with a custom OData filter
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="filter">OData filter to apply to retrieval of the users from the Microsoft Graph</param>
        /// <param name="orderby">OData orderby instruction</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <returns>List with User objects</returns>
        public static List<Model.User> ListUsers(string accessToken, string filter, string orderby, int startIndex = 0, int endIndex = 999, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            List<Model.User> result = null;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    List<Model.User> users = new List<Model.User>();

                    var graphClient = GraphUtility.CreateGraphClient(accessToken, retryCount, delay);

                    var pagedUsers = await graphClient.Users
                        .Request()
                        .Filter(filter)
                        .OrderBy(orderby)
                        .GetAsync();

                    int pageCount = 0;
                    int currentIndex = 0;

                    while (true)
                    {
                        pageCount++;

                        foreach (var u in pagedUsers)
                        {
                            currentIndex++;

                            if (currentIndex >= startIndex)
                            {
                                var user = new Model.User
                                {
                                    Id = Guid.TryParse(u.Id, out Guid idGuid) ? (Guid?) idGuid : null,
                                    DisplayName = u.DisplayName,
                                    GivenName = u.GivenName,
                                    JobTitle = u.JobTitle,
                                    MobilePhone = u.MobilePhone,
                                    OfficeLocation = u.OfficeLocation,
                                    PreferredLanguage = u.PreferredLanguage,
                                    Surname = u.Surname,
                                    UserPrincipalName = u.UserPrincipalName
                                };

                                users.Add(user);
                            }
                        }

                        if (pagedUsers.NextPageRequest != null && users.Count < endIndex)
                        {
                            pagedUsers = await pagedUsers.NextPageRequest.GetAsync();
                        }
                        else
                        {
                            break;
                        }
                    }

                    return users;
                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return result;
        }
    }
}
