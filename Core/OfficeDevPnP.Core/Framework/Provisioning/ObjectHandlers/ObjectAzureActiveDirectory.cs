#if !ONPREMISES
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Linq;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Utilities;
using System.Net;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Web;
using System.Net.Http;
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.Provisioning.Model.AzureActiveDirectory;
using OfficeDevPnP.Core.Utilities.Graph;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using System.IO;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Object Handler to manage Microsoft AAD stuff
    /// </summary>
    internal class ObjectAzureActiveDirectory : ObjectHierarchyHandlerBase
    {
        public override string Name => "AzureActiveDirectory ";

        /// <summary>
        /// Creates a User in AAD and configures password and services
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="user">The User to create</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>The ID of the User</returns>
        private object CreateOrUpdateUser(PnPMonitoredScope scope, TokenParser parser, Model.AzureActiveDirectory.User user, string accessToken)
        {
            var content = PrepareUserRequestContent(user, parser);

            var userId = GraphHelper.CreateOrUpdateGraphObject(scope,
                HttpMethodVerb.POST,
                $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users",
                content,
                HttpHelper.JsonContentType,
                accessToken,
                "ObjectConflict",
                CoreResources.Provisioning_ObjectHandlers_AAD_User_AlreadyExists,
                "userPrincipalName",
                parser.ParseString(user.UserPrincipalName),
                CoreResources.Provisioning_ObjectHandlers_AAD_User_ProvisioningError,
                canPatch: true);

            return (userId);
        }

        /// <summary>
        /// Prepares the object to serialize as JSON for adding/updating a User object
        /// </summary>
        /// <param name="user">The source User object</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <returns>The User object to serialize as JSON</returns>
        private object PrepareUserRequestContent(Model.AzureActiveDirectory.User user, TokenParser parser)
        {
            var content = new
            {
                accountEnabled = user.AccountEnabled,
                displayName = parser.ParseString(user.DisplayName),
                mailNickname = parser.ParseString(user.MailNickname),
                userPrincipalName = parser.ParseString(user.UserPrincipalName),
                givenName = parser.ParseString(user.GivenName),
                surname = parser.ParseString(user.Surname),
                jobTitle = parser.ParseString(user.JobTitle),
                mobilePhone = parser.ParseString(user.MobilePhone),
                officeLocation = parser.ParseString(user.OfficeLocation),
                preferredLanguage = parser.ParseString(user.PreferredLanguage),
                userType = "Member",
                usageLocation = parser.ParseString(user.UsageLocation),
                passwordPolicies = parser.ParseString(user.PasswordPolicies),
                passwordProfile = new
                {
                    forceChangePasswordNextSignIn = user.PasswordProfile.ForceChangePasswordNextSignIn,
                    forceChangePasswordNextSignInWithMfa = user.PasswordProfile.ForceChangePasswordNextSignInWithMfa,
                    password = EncryptionUtility.ToInsecureString(user.PasswordProfile.Password),
                }
            };

            return (content);
        }

        /// <summary>
        /// Manages User licenses with delta handling
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="userId">The ID of the target user</param>
        /// <param name="licenses">The Licenses to manage</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        private void ManageUserLicenses(PnPMonitoredScope scope, object userId, UserLicenseCollection licenses, string accessToken)
        {
            // Get the currently assigned licenses
            var jsoncurrentLicenses = HttpHelper.MakeGetRequestForString(
                $"{GraphHelper.MicrosoftGraphBaseURI}beta/users/{userId}", accessToken);
            var currentLicenses = JsonConvert.DeserializeAnonymousType(jsoncurrentLicenses, new
            {
                assignedLicenses = new[]
                {
                    new {
                        disabledPlans = new[]
                        {
                            Guid.Empty
                        },
                        skuId = Guid.Empty
                    }
                }
            });

            // Manage the license to remove
            var removeLicenses = new List<Guid>();
            foreach (var l in currentLicenses.assignedLicenses)
            {
                // If the already assigned license is not in the list of new licenses
                if (!licenses.Any(lic => Guid.Parse(lic.SkuId) == l.skuId))
                {
                    // We need to remove it
                    removeLicenses.Add(l.skuId);
                }
            }

            // Prepare the new request to update assigned licenses
            var assigneLicenseBody = new
            {
                addLicenses = from l in licenses
                              select new
                              {
                                  skuId = Guid.Parse(l.SkuId),
                                  disabledPlans = l.DisabledPlans != null ?
                                    (from d in l.DisabledPlans
                                     select Guid.Parse(d)).ToArray() : new Guid[] { }
                              },
                removeLicenses = (from r in removeLicenses
                                  select r).ToArray()
            };
            HttpHelper.MakePostRequest(
                $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{userId}/assignLicense",
                assigneLicenseBody, HttpHelper.JsonContentType, accessToken);
        }

        /// <summary>
        /// Synchronizes User's Photo
        /// </summary>
        /// <param name="scope">The PnP Provisioning Scope</param>
        /// <param name="parser">The PnP Token Parser</param>
        /// <param name="connector">The PnP file connector</param>
        /// <param name="user">The target User</param>
        /// <param name="userId">The ID of the target User</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token</param>
        /// <returns>Whether the Photo has been updated or not</returns>
        private static bool SetUserPhoto(PnPMonitoredScope scope, TokenParser parser, FileConnectorBase connector, Model.AzureActiveDirectory.User user, string userId, string accessToken)
        {
            Boolean result = false;

            if (!String.IsNullOrEmpty(user.ProfilePhoto) && connector != null)
            {
                var photoPath = parser.ParseString(user.ProfilePhoto);
                var photoBytes = ConnectorFileHelper.GetFileBytes(connector, user.ProfilePhoto);

                using (var mem = new MemoryStream())
                {
                    mem.Write(photoBytes, 0, photoBytes.Length);
                    mem.Position = 0;

                    HttpHelper.MakePostRequest(
                        $"{GraphHelper.MicrosoftGraphBaseURI}v1.0/users/{userId}/photo/$value",
                        mem, "image/jpeg", accessToken);
                }
            }

            return (result);
        }

        #region PnP Provisioning Engine infrastructural code

        public override bool WillProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration)
        {
#if !ONPREMISES
            if (!_willProvision.HasValue)
            {
                _willProvision = hierarchy.AzureActiveDirectory?.Users?.Any();
            }
#else
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
#endif
            return _willProvision.Value;
        }

        public override bool WillExtract(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ExtractConfiguration configuration)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }

        public override TokenParser ProvisionObjects(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser, ApplyConfiguration configuration)
        {
#if !ONPREMISES
            using (var scope = new PnPMonitoredScope(Name))
            {
                // Prepare a method global variable to store the Access Token
                String accessToken = null;

                // - Teams based on JSON templates
                var users = hierarchy.AzureActiveDirectory?.Users;
                if (users != null && users.Any())
                {
                    foreach (var u in users)
                    {
                        // Get a fresh Access Token for every request
                        accessToken = PnPProvisioningContext.Current.AcquireToken(GraphHelper.MicrosoftGraphBaseURI, "User.ReadWrite.All");

                        // Creates or updates the User starting from the provisioning template definition
                        var userId = CreateOrUpdateUser(scope, parser, u, accessToken);

                        // If the user got created
                        if (userId != null)
                        {
                            if (u.Licenses != null && u.Licenses.Count > 0)
                            {
                                // Manage the licensing settings
                                ManageUserLicenses(scope, userId, u.Licenses, accessToken);
                            }

                            // So far the User's photo cannot be set if we don't have an already existing mailbox
                            // SetUserPhoto(scope, parser, hierarchy.Connector, u, (String)userId, accessToken);
                        }
                    }
                }
            }
#endif

            return parser;
        }

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ExtractConfiguration configuration)
        {
            // So far, no extraction
            return hierarchy;
        }

        #endregion
    }
}
#endif