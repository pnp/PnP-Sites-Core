using Microsoft.SharePoint.ApplicationPages.ClientPickerQuery;
using OfficeDevPnP.Core.Utilities.Async;
using System;
using System.Threading.Tasks;

#if !ONPREMISES
namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Defines the options for Sharing Document
    /// </summary>
    public enum ExternalSharingDocumentOption
    {
        /// <summary>
        /// Provides the option to edit the document
        /// </summary>
        Edit,
        /// <summary>
        /// Provides the option to view the document
        /// </summary>
        View
    }

    /// <summary>
    /// Defines the options for Sharing Site
    /// </summary>
    public enum ExternalSharingSiteOption
    {
        /// <summary>
        /// Provides sharing to AssociatedOwnerGroup 
        /// </summary>
        Owner,
        /// <summary>
        /// Provides sharing to AssociatedMemberGroup 
        /// </summary>
        Edit,
        /// <summary>
        /// Provides sharing to AssociatedVisitorGroup 
        /// </summary>
        View
    }

    /// <summary>
    /// This class holds the methods for sharing and unsharing of the document and the site. 
    /// </summary>
    public static partial class ExternalSharingExtensions
    {
        /// <summary>
        /// Can be used to get needed people picker search result value for given email account. 
        /// See <a href="https://msdn.microsoft.com/en-us/library/office/jj179690.aspx">MSDN</a>
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="emailAddress">Email address to be used as the query parameter. Should be pointing to unique person which is then searched using people picker capability programatically.</param>
        /// <returns>Resolves people picker value which can be used for sharing objects in the SharePoint site</returns>
        public static string ResolvePeoplePickerValueForEmail(this Web web, string emailAddress)
        {
            return Task.Run(() => ResolvePeoplePickerValueForEmailImplementation(web, emailAddress)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Can be used to get needed people picker search result value for given email account. 
        /// See <a href="https://msdn.microsoft.com/en-us/library/office/jj179690.aspx">MSDN</a>
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="emailAddress">Email address to be used as the query parameter. Should be pointing to unique person which is then searched using people picker capability programatically.</param>
        /// <returns>Resolves people picker value which can be used for sharing objects in the SharePoint site</returns>
        public static async Task<string> ResolvePeoplePickerValueForEmailAsync(this Web web, string emailAddress)
        {
            await new SynchronizationContextRemover();
            return await ResolvePeoplePickerValueForEmailImplementation(web, emailAddress);
        }

        private static async Task<string> ResolvePeoplePickerValueForEmailImplementation(Web web, string emailAddress)
        {
            var param = new ClientPeoplePickerQueryParameters
            {
                PrincipalSource = Utilities.PrincipalSource.All,
                PrincipalType = Utilities.PrincipalType.All,
                MaximumEntitySuggestions = 30,
                QueryString = emailAddress,
                AllowEmailAddresses = true,
                AllowOnlyEmailAddresses = false,
                AllUrlZones = false,
                ForceClaims = false,
                Required = true,
                SharePointGroupID = 0,
                UrlZone = 0,
                UrlZoneSpecified = false
            };

            // Resolve people picker value based on email
            var ret = ClientPeoplePickerWebServiceInterface.ClientPeoplePickerResolveUser(web.Context, param);
            await web.Context.ExecuteQueryRetryAsync();

            // Return people picker return value in right format
            return $"[{ret.Value}]";
        }

        /// <summary>
        /// Creates anonymous link to given document.
        /// See <a href="https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.web.createanonymouslink.aspx">MSDN</a>
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="shareOption">Type of the link to be created - View or Edit</param>
        /// <returns>Anonymous URL to the file as string</returns>
        public static string CreateAnonymousLinkForDocument(this Web web, string urlToDocument, ExternalSharingDocumentOption shareOption)
        {
            return Task.Run(() => CreateAnonymousLinkForDocumentImplementation(web, urlToDocument, shareOption)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Creates anonymous link to given document.
        /// See <a href="https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.web.createanonymouslink.aspx">MSDN</a>
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="shareOption">Type of the link to be created - View or Edit</param>
        /// <returns>Anonymous URL to the file as string</returns>
        public static async Task<string> CreateAnonymousLinkForDocumentAsync(this Web web, string urlToDocument, ExternalSharingDocumentOption shareOption)
        {
            await new SynchronizationContextRemover();
            return await CreateAnonymousLinkForDocumentImplementation(web, urlToDocument, shareOption);
        }

        private static async Task<string> CreateAnonymousLinkForDocumentImplementation(Web web, string urlToDocument, ExternalSharingDocumentOption shareOption)
        {
            bool isEditLink = true;
            switch (shareOption)
            {
                case ExternalSharingDocumentOption.Edit:
                    isEditLink = true;
                    break;
                case ExternalSharingDocumentOption.View:
                    isEditLink = false;
                    break;
                default:
                    break;
            }
            var result = Web.CreateAnonymousLink(web.Context, urlToDocument, isEditLink);
            await web.Context.ExecuteQueryRetryAsync();

            // return anonymous link to caller
            return result.Value;
        }

        /// <summary>
        /// Creates anonymous link to the given document with automatic expiration time.
        /// See <a href="https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.web.createanonymouslinkwithexpiration.aspx">MSDN</a>
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="shareOption">Type of the link to be created - View or Edit</param>
        /// <param name="expireTime">Date time for link expiration - will be converted to ISO 8601 format automatically</param>
        /// <returns>Anonymous URL to the file as string</returns>
        public static string CreateAnonymousLinkWithExpirationForDocument(this Web web, string urlToDocument, ExternalSharingDocumentOption shareOption, DateTime expireTime)
        {
            return Task.Run(() => CreateAnonymousLinkWithExpirationForDocumentImplementation(web, urlToDocument, shareOption, expireTime)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Creates anonymous link to the given document with automatic expiration time.
        /// See <a href="https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.web.createanonymouslinkwithexpiration.aspx">MSDN</a>
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="shareOption">Type of the link to be created - View or Edit</param>
        /// <param name="expireTime">Date time for link expiration - will be converted to ISO 8601 format automatically</param>
        /// <returns>Anonymous URL to the file as string</returns>
        public static async Task<string> CreateAnonymousLinkWithExpirationForDocumentAsync(this Web web, string urlToDocument, ExternalSharingDocumentOption shareOption, DateTime expireTime)
        {
            await new SynchronizationContextRemover();
            return await CreateAnonymousLinkWithExpirationForDocumentImplementation(web, urlToDocument, shareOption, expireTime);
        }

        private static async Task<string> CreateAnonymousLinkWithExpirationForDocumentImplementation(Web web, string urlToDocument, ExternalSharingDocumentOption shareOption, DateTime expireTime)
        {
            // If null given as expiration, there will not be automatic expiration time
            var expirationTimeAsString = expireTime.ToString("s", System.Globalization.CultureInfo.InvariantCulture);

            bool isEditLink = true;
            switch (shareOption)
            {
                case ExternalSharingDocumentOption.Edit:
                    isEditLink = true;
                    break;
                case ExternalSharingDocumentOption.View:
                    isEditLink = false;
                    break;
                default:
                    break;
            }

            // Get the link
            var result = Web.CreateAnonymousLinkWithExpiration(web.Context, urlToDocument, isEditLink, expirationTimeAsString);
            await web.Context.ExecuteQueryRetryAsync();

            // Return anonymous link to caller
            return result.Value;
        }

        /// <summary>
        /// Abstracted methid for sharing documents just with given email address. 
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="targetEmailToShare">Email address for the person to whom the document will be shared</param>
        /// <param name="shareOption">View or Edit option</param>
        /// <param name="sendEmail">Send email or not</param>
        /// <param name="emailBody">Text attached to the email sent for the person to whom the document is shared</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <see cref="ShareDocument(Web, string, string, ExternalSharingDocumentOption, bool, string, bool)"/>
        /// <returns>A SharingResult object</returns>
        public static SharingResult ShareDocument(this Web web, string urlToDocument,
                                                string targetEmailToShare, ExternalSharingDocumentOption shareOption,
                                                bool sendEmail = true, string emailBody = "Document shared",
                                                bool useSimplifiedRoles = true)
        {
            return Task.Run(() => ShareDocumentImplementation(web, urlToDocument, targetEmailToShare, shareOption, sendEmail, emailBody, useSimplifiedRoles)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Abstracted methid for sharing documents just with given email address. 
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="targetEmailToShare">Email address for the person to whom the document will be shared</param>
        /// <param name="shareOption">View or Edit option</param>
        /// <param name="sendEmail">Send email or not</param>
        /// <param name="emailBody">Text attached to the email sent for the person to whom the document is shared</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <see cref="ShareDocumentAsync(Web, string, string, ExternalSharingDocumentOption, bool, string, bool)"/>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> ShareDocumentAsync(this Web web, string urlToDocument,
                                                string targetEmailToShare, ExternalSharingDocumentOption shareOption,
                                                bool sendEmail = true, string emailBody = "Document shared",
                                                bool useSimplifiedRoles = true)
        {
            await new SynchronizationContextRemover();
            return await ShareDocumentImplementation(web, urlToDocument, targetEmailToShare, shareOption, sendEmail, emailBody, useSimplifiedRoles);
        }

        private static async Task<SharingResult> ShareDocumentImplementation(Web web, string urlToDocument, string targetEmailToShare, ExternalSharingDocumentOption shareOption, bool sendEmail, string emailBody, bool useSimplifiedRoles)
        {
            // Resolve people picker value for given email
            string peoplePickerInput = await ResolvePeoplePickerValueForEmailImplementation(web, targetEmailToShare);
            // Share document for user
            return await ShareDocumentWithPeoplePickerValueImplementation(web, urlToDocument, peoplePickerInput, shareOption, sendEmail, emailBody, useSimplifiedRoles);
        }

        /// <summary>
        /// Share document with complex JSON string value.
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="peoplePickerInput">People picker JSON string value containing the target person information</param>
        /// <param name="shareOption">View or Edit option</param>
        /// <param name="sendEmail">Send email or not</param>
        /// <param name="emailBody">Text attached to the email sent for the person to whom the document is shared</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <returns>A SharingResult object</returns>
        public static SharingResult ShareDocumentWithPeoplePickerValue(this Web web, string urlToDocument, string peoplePickerInput,
                                        ExternalSharingDocumentOption shareOption, bool sendEmail = true,
                                        string emailBody = "Document shared for you.", bool useSimplifiedRoles = true)
        {
            return Task.Run(() => ShareDocumentWithPeoplePickerValueImplementation(web, urlToDocument, peoplePickerInput, shareOption, sendEmail, emailBody, useSimplifiedRoles)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Share document with complex JSON string value.
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="peoplePickerInput">People picker JSON string value containing the target person information</param>
        /// <param name="shareOption">View or Edit option</param>
        /// <param name="sendEmail">Send email or not</param>
        /// <param name="emailBody">Text attached to the email sent for the person to whom the document is shared</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> ShareDocumentWithPeoplePickerValueAsync(this Web web, string urlToDocument, string peoplePickerInput,
                                        ExternalSharingDocumentOption shareOption, bool sendEmail = true,
                                        string emailBody = "Document shared for you.", bool useSimplifiedRoles = true)
        {
            await new SynchronizationContextRemover();
            return await ShareDocumentWithPeoplePickerValueImplementation(web, urlToDocument, peoplePickerInput, shareOption, sendEmail, emailBody, useSimplifiedRoles);
        }

        private static async Task<SharingResult> ShareDocumentWithPeoplePickerValueImplementation(Web web, string urlToDocument, string peoplePickerInput, ExternalSharingDocumentOption shareOption, bool sendEmail, string emailBody, bool useSimplifiedRoles)
        {
            var groupId = 0;            // Set groupId to 0 for external share
            var propageAcl = false;    // Not relevant for external accounts
            string emailSubject = null; // Not relevant, since we can't change subject
            var includedAnonymousLinkInEmail = false;  // Check if this has any meaning in first place

            // Set role value accordingly based on requested share option - These are constant in the server side code.
            var roleValue = "";
            switch (shareOption)
            {
                case ExternalSharingDocumentOption.Edit:
                    roleValue = "role:1073741827";
                    break;
                default:
                    // Use this for other options - Means View permission
                    roleValue = "role:1073741826";
                    break;
            }

            // Share the document, send email and return the result value
            var result = Web.ShareObject(web.Context, urlToDocument,
                                                        peoplePickerInput, roleValue, groupId, propageAcl,
                                                        sendEmail, includedAnonymousLinkInEmail, emailSubject,
                                                        emailBody, useSimplifiedRoles);

            web.Context.Load(result);
            await web.Context.ExecuteQueryRetryAsync();
            return result;
        }

        /// <summary>
        /// Can be used to programatically to unshare any document with the document URL.
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <returns>A SharingResult object</returns>
        public static SharingResult UnshareDocument(this Web web, string urlToDocument)
        {
            return Task.Run(() => UnshareDocumentImplementation(web, urlToDocument)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Can be used to programatically to unshare any document with the document URL.
        /// </summary>
        /// <param name="web">Web for the context used for people picker search</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> UnshareDocumentAsync(this Web web, string urlToDocument)
        {
            await new SynchronizationContextRemover();
            return await UnshareDocumentImplementation(web, urlToDocument);
        }

        private static async Task<SharingResult> UnshareDocumentImplementation(Web web, string urlToDocument)
        {
            var result = Web.UnshareObject(web.Context, urlToDocument);
            web.Context.Load(result);
            await web.Context.ExecuteQueryRetryAsync();

            // Return the results
            return result;
        }

        /// <summary>
        /// Get current sharing settings for document and load list of users it has been shared automatically.
        /// </summary>
        /// <param name="web">Web for the context</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="useSimplifiedPolicies">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View).</param>
        /// <returns>A ObjectSharingSettings object</returns>
        public static ObjectSharingSettings GetObjectSharingSettingsForDocument(this Web web, string urlToDocument, bool useSimplifiedPolicies = true)
        {
            return Task.Run(() => GetObjectSharingSettingsForDocumentImplementation(web, urlToDocument, useSimplifiedPolicies)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Get current sharing settings for document and load list of users it has been shared automatically.
        /// </summary>
        /// <param name="web">Web for the context</param>
        /// <param name="urlToDocument">Full URL to the file which is shared</param>
        /// <param name="useSimplifiedPolicies">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View).</param>
        /// <returns>A ObjectSharingSettings object</returns>
        public static async Task<ObjectSharingSettings> GetObjectSharingSettingsForDocumentAsync(this Web web, string urlToDocument, bool useSimplifiedPolicies = true)
        {
            await new SynchronizationContextRemover();
            return await GetObjectSharingSettingsForDocumentImplementation(web, urlToDocument, useSimplifiedPolicies);
        }

        private static async Task<ObjectSharingSettings> GetObjectSharingSettingsForDocumentImplementation(Web web, string urlToDocument, bool useSimplifiedPolicies)
        {
            // Group value for this query is always 0.
            var info = Web.GetObjectSharingSettings(web.Context, urlToDocument, 0, useSimplifiedPolicies);
            web.Context.Load(info);
            web.Context.Load(info.ObjectSharingInformation);
            web.Context.Load(info.ObjectSharingInformation.SharedWithUsersCollection);
            await web.Context.ExecuteQueryRetryAsync();

            return info;
        }

        /// <summary>
        /// Get current sharing settings for site and load list of users it has been shared automatically.
        /// </summary>
        /// <param name="web">Web for the context</param>
        /// <param name="useSimplifiedPolicies"></param>
        /// <returns>A ObjectSharingSettings object</returns>
        public static ObjectSharingSettings GetObjectSharingSettingsForSite(this Web web, bool useSimplifiedPolicies = true)
        {
            return Task.Run(() => GetObjectSharingSettingsForSiteImplementation(web, useSimplifiedPolicies)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Get current sharing settings for site and load list of users it has been shared automatically.
        /// </summary>
        /// <param name="web">Web for the context</param>
        /// <param name="useSimplifiedPolicies"></param>
        /// <returns>A ObjectSharingSettings object</returns>
        public static async Task<ObjectSharingSettings> GetObjectSharingSettingsForSiteAsync(this Web web, bool useSimplifiedPolicies = true)
        {
            await new SynchronizationContextRemover();
            return await GetObjectSharingSettingsForSiteImplementation(web, useSimplifiedPolicies);
        }

        private static async Task<ObjectSharingSettings> GetObjectSharingSettingsForSiteImplementation(Web web, bool useSimplifiedPolicies)
        {
            // Ensure that URL exists
            if (!web.IsObjectPropertyInstantiated("Url"))
            {
                web.Context.Load(web, w => w.Url);
                await web.Context.ExecuteQueryRetryAsync();
            }

            ObjectSharingSettings info = Web.GetObjectSharingSettings(web.Context, web.Url, 0, useSimplifiedPolicies);
            web.Context.Load(info);
            web.Context.Load(info.ObjectSharingInformation);
            web.Context.Load(info.ObjectSharingInformation.SharedWithUsersCollection);
            await web.Context.ExecuteQueryRetryAsync();

            return info;
        }

        /// <summary>
        /// Invites an external user as a group member
        /// </summary>
        /// <param name="group">Group to add the user to</param>
        /// <param name="email">The email address of the external user</param>
        /// <param name="sendEmail">Should we send an email to the given address</param>
        /// <param name="emailBody">Text to be added to the email</param>
        /// <returns>A SharingResult object</returns>
        public static SharingResult InviteExternalUser(this Group group, string email, bool sendEmail = true,
            string emailBody = "Site shared with you.")
        {
            return Task.Run(() => InviteExternalUserImplementation(group, email, sendEmail, emailBody)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Invites an external user as a group member
        /// </summary>
        /// <param name="group">Group to add the user to</param>
        /// <param name="email">The email address of the external user</param>
        /// <param name="sendEmail">Should we send an email to the given address</param>
        /// <param name="emailBody">Text to be added to the email</param>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> InviteExternalUserAsync(this Group group, string email, bool sendEmail = true,
            string emailBody = "Site shared with you.")
        {
            await new SynchronizationContextRemover();
            return await InviteExternalUserImplementation(group, email, sendEmail, emailBody);
        }

        private static async Task<SharingResult> InviteExternalUserImplementation(Group group, string email, bool sendEmail, string emailBody)
        {
            var web = (group.Context as ClientContext).Web;

            return await ShareSiteImplementation(web, email, group, sendEmail, emailBody);
        }

        /// <summary>
        /// Share site for a person using just email. Will resolve needed people picker JSON value automatically.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="email">Email of the person to whom site should be shared.</param>
        /// <param name="group">Group to invite the external user to</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <returns>A SharingResult object</returns>
        public static SharingResult ShareSite(this Web web, string email,
            Group group, bool sendEmail = true, string emailBody = "Site shared for you.")
        {
            return Task.Run(() => ShareSiteImplementation(web, email, group, sendEmail, emailBody)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Share site for a person using just email. Will resolve needed people picker JSON value automatically.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="email">Email of the person to whom site should be shared.</param>
        /// <param name="group">Group to invite the external user to</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> ShareSiteAsync(this Web web, string email,
            Group group, bool sendEmail = true, string emailBody = "Site shared for you.")
        {
            await new SynchronizationContextRemover();
            return await ShareSiteImplementation(web, email, group, sendEmail, emailBody);
        }

        private static async Task<SharingResult> ShareSiteImplementation(Web web, string email, Group group, bool sendEmail, string emailBody)
        {
            var peoplePickerValue = await ResolvePeoplePickerValueForEmailImplementation(web, email);
            return await ShareSiteWithPeoplePickerValueImplementation(web, peoplePickerValue, group, sendEmail, emailBody);
        }

        /// <summary>
        /// Share site for a person using just email. Will resolve needed people picker JSON value automatically.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="email">Email of the person to whom site should be shared.</param>
        /// <param name="shareOption">Sharing style - View, Edit, Owner</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <returns>A SharingResult object</returns>
        public static SharingResult ShareSite(this Web web, string email,
                                                ExternalSharingSiteOption shareOption, bool sendEmail = true,
                                                string emailBody = "Site shared for you.", bool useSimplifiedRoles = true)
        {
            return Task.Run(() => ShareSiteImplementation(web, email, shareOption, sendEmail, emailBody, useSimplifiedRoles)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Share site for a person using just email. Will resolve needed people picker JSON value automatically.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="email">Email of the person to whom site should be shared.</param>
        /// <param name="shareOption">Sharing style - View, Edit, Owner</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> ShareSiteAsync(this Web web, string email,
                                                ExternalSharingSiteOption shareOption, bool sendEmail = true,
                                                string emailBody = "Site shared for you.", bool useSimplifiedRoles = true)
        {
            await new SynchronizationContextRemover();
            return await ShareSiteImplementation(web, email, shareOption, sendEmail, emailBody, useSimplifiedRoles);
        }

        private static async Task<SharingResult> ShareSiteImplementation(Web web, string email, ExternalSharingSiteOption shareOption, bool sendEmail, string emailBody, bool useSimplifiedRoles)
        {
            // Solve people picker value for email address
            string peoplePickerValue = await ResolvePeoplePickerValueForEmailImplementation(web, email);

            // Share with the people picker value
            return await ShareSiteWithPeoplePickerValueImplementation(web, peoplePickerValue, shareOption, sendEmail, emailBody, useSimplifiedRoles);
        }

        /// <summary>
        /// Share site for a person using complex JSON object for people picker value.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="peoplePickerInput">JSON object with the people picker value</param>
        /// <param name="shareOption">Sharing style - View, Edit, Owner</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <returns>A SharingResult object</returns>
        public static SharingResult ShareSiteWithPeoplePickerValue(this Web web, string peoplePickerInput,
                                                                    ExternalSharingSiteOption shareOption,
                                                                    bool sendEmail = true, string emailBody = "Site shared for you.",
                                                                    bool useSimplifiedRoles = true)
        {
            return Task.Run(() => ShareSiteWithPeoplePickerValueImplementation(web, peoplePickerInput, shareOption, sendEmail, emailBody, useSimplifiedRoles)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Share site for a person using complex JSON object for people picker value.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="peoplePickerInput">JSON object with the people picker value</param>
        /// <param name="shareOption">Sharing style - View, Edit, Owner</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <param name="useSimplifiedRoles">Boolean value indicating whether to use the SharePoint simplified roles (Edit, View)</param>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> ShareSiteWithPeoplePickerValueAsync(this Web web, string peoplePickerInput,
                                                                    ExternalSharingSiteOption shareOption,
                                                                    bool sendEmail = true, string emailBody = "Site shared for you.",
                                                                    bool useSimplifiedRoles = true)
        {
            await new SynchronizationContextRemover();
            return await ShareSiteWithPeoplePickerValueImplementation(web, peoplePickerInput, shareOption, sendEmail, emailBody, useSimplifiedRoles);
        }

        private static async Task<SharingResult> ShareSiteWithPeoplePickerValueImplementation(Web web, string peoplePickerInput, ExternalSharingSiteOption shareOption, bool sendEmail, string emailBody, bool useSimplifiedRoles)
        {
            // Solve the group id for the shared option based on default groups
            var groupId = await SolveGroupIdToShareAsync(web, shareOption);
            string roleValue = $"group:{groupId}"; // Right permission setup

            web.EnsureProperty(w => w.Url);

            // Set default settings for site sharing
            var propagateAcl = false; // Not relevant for external accounts
            var includedAnonymousLinkInEmail = false; // Not when site is shared

            var result = Web.ShareObject(web.Context, web.Url, peoplePickerInput,
                                                        roleValue, 0, propagateAcl,
                                                        sendEmail, includedAnonymousLinkInEmail, null,
                                                        emailBody, useSimplifiedRoles);
            web.Context.Load(result);
            await web.Context.ExecuteQueryRetryAsync();
            return result;
        }

        /// <summary>
        /// Share site for a person using complex JSON object for people picker value.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="peoplePickerInput">JSON object with the people picker value</param>
        /// <param name="group">The group to invite the user to</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <returns>A SharingResult object</returns>
        public static SharingResult ShareSiteWithPeoplePickerValue(this Web web, string peoplePickerInput,
                                                                    Group group,
                                                                    bool sendEmail = true, string emailBody = "Site shared for you.")
        {
            return Task.Run(() => ShareSiteWithPeoplePickerValueImplementation(web, peoplePickerInput, group, sendEmail, emailBody)).GetAwaiter().GetResult();
        }
        /// <summary>
        /// Share site for a person using complex JSON object for people picker value.
        /// </summary>
        /// <param name="web">Web for the context of the site to be shared.</param>
        /// <param name="peoplePickerInput">JSON object with the people picker value</param>
        /// <param name="group">The group to invite the user to</param>
        /// <param name="sendEmail">Should we send email for the given address.</param>
        /// <param name="emailBody">Text to be added on share email sent to receiver.</param>
        /// <returns>A SharingResult object</returns>
        public static async Task<SharingResult> ShareSiteWithPeoplePickerValueAsync(this Web web, string peoplePickerInput,
                                                                    Group group,
                                                                    bool sendEmail = true, string emailBody = "Site shared for you.")
        {
            await new SynchronizationContextRemover();
            return await ShareSiteWithPeoplePickerValueImplementation(web, peoplePickerInput, group, sendEmail, emailBody);
        }

        private static async Task<SharingResult> ShareSiteWithPeoplePickerValueImplementation(Web web, string peoplePickerInput, Group group, bool sendEmail, string emailBody)
        {
            // Solve the group id for the shared option based on default groups
            var groupId = group.Id;
            string roleValue = $"group:{groupId}"; // Right permission setup

            web.EnsureProperty(w => w.Url);

            var result = Web.ShareObject(web.Context, web.Url, peoplePickerInput, roleValue, groupId, false,
                sendEmail, false, null, emailBody, false);

            web.Context.Load(result);
            await web.Context.ExecuteQueryRetryAsync();
            return result;
        }

        /// <summary>
        /// Used to solve right group ID to assign user into - used for the site level sharing.
        /// </summary>
        /// <param name="web">Web to be shared externally</param>
        /// <param name="shareOption">Permissions to be given for the external user</param>
        /// <returns>group ID</returns>
        private static async Task<int> SolveGroupIdToShareAsync(Web web, ExternalSharingSiteOption shareOption)
        {
            Group group = null;
            switch (shareOption)
            {
                case ExternalSharingSiteOption.Owner:
                    group = web.AssociatedOwnerGroup;
                    break;
                case ExternalSharingSiteOption.Edit:
                    group = web.AssociatedMemberGroup;
                    break;
                case ExternalSharingSiteOption.View:
                    group = web.AssociatedVisitorGroup;
                    break;
                default:
                    group = web.AssociatedVisitorGroup;
                    break;
            }
            // Load right group
            web.Context.Load(group);
            await web.Context.ExecuteQueryRetryAsync();
            // Return group ID
            return group.Id;
        }
    }
}
#endif