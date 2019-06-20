using Microsoft.Graph;
using OfficeDevPnP.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Graph
{
    /// <summary>
    /// Utility class to perform Graph operations.
    /// </summary>
    public static class GraphUtility
    {
        private const int defaultRetryCount = 10;
        private const int defaultDelay = 500;

        /// <summary>
        ///  Creates a new GraphServiceClient instance using a custom PnPHttpProvider
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to configure the HTTP bearer Authorization Header</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request.</param>
        /// <returns></returns>
        public static GraphServiceClient CreateGraphClient(String accessToken, int retryCount = defaultRetryCount, int delay = defaultDelay)
        {
            // Creates a new GraphServiceClient instance using a custom PnPHttpProvider
            // which natively supports retry logic for throttled requests
            // Default are 10 retries with a base delay of 500ms
            var result = new GraphServiceClient(new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            if (!String.IsNullOrEmpty(accessToken))
                            {
                                // Configure the HTTP bearer Authorization Header
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                            }
                        }), new PnPHttpProvider(retryCount, delay));

            return (result);
        }

        /// <summary>
        /// This method sends an Azure guest user invitation to the provided email address.
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="guestUserEmail">Email of the user to whom the invite must be sent</param>
        /// <param name="redirectUri">URL where the user will be redirected after the invite is accepted.</param>
        /// <param name="customizedMessage">Customized email message to be sent in the invitation email.</param>
        /// <param name="guestUserDisplayName">Display name of the Guest user.</param>
        /// <returns></returns>
        public static Invitation InviteGuestUser(string accessToken, string guestUserEmail, string redirectUri, string customizedMessage = "", string guestUserDisplayName = "")
        {
            Invitation inviteUserResponse = null;

            try
            {
                Invitation invite = new Invitation();
                invite.InvitedUserEmailAddress = guestUserEmail;
                if (!string.IsNullOrWhiteSpace(guestUserDisplayName))
                {
                    invite.InvitedUserDisplayName = guestUserDisplayName;
                }
                invite.InviteRedirectUrl = redirectUri;
                invite.SendInvitationMessage = true;

                // Form the invite email message body
                if (!string.IsNullOrWhiteSpace(customizedMessage))
                {
                    InvitedUserMessageInfo inviteMsgInfo = new InvitedUserMessageInfo();
                    inviteMsgInfo.CustomizedMessageBody = customizedMessage;
                    invite.InvitedUserMessageInfo = inviteMsgInfo;
                }

                // Create the graph client and send the invitation.
                GraphServiceClient graphClient = CreateGraphClient(accessToken);
                inviteUserResponse = graphClient.Invitations.Request().AddAsync(invite).Result;
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return inviteUserResponse;
        }
    }
}
