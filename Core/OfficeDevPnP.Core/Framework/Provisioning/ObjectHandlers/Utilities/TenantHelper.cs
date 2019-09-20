#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantAdministration.Internal;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.ALM;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Security.Cryptography;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    internal static class TenantHelper
    {
        internal static string appExistsQuery = @"<View>
<Query>
   <Where>
      <Eq>
         <FieldRef Name='AppPackageHash' />
         <Value Type='Text'>{0}</Value>
      </Eq>
   </Where>
</Query>
</View>";
        public static TokenParser ProcessApps(Tenant tenant, ProvisioningTenant provisioningTenant, FileConnectorBase connector, TokenParser parser, PnPMonitoredScope scope, ProvisioningTemplateApplyingInformation applyingInformation, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.AppCatalog != null && provisioningTenant.AppCatalog.Packages.Count > 0)
            {
                var rootSiteUrl = tenant.GetRootSiteUrl();
                tenant.Context.ExecuteQueryRetry();
                using (var context = ((ClientContext)tenant.Context).Clone(rootSiteUrl.Value, applyingInformation.AccessTokens))
                {

                    var web = context.Web;

                    Uri appCatalogUri = null;

                    try
                    {
                        appCatalogUri = web.GetAppCatalog();
                    }
                    catch (System.Net.WebException ex)
                    {
                        if (ex.Response != null)
                        {
                            var httpResponse = ex.Response as System.Net.HttpWebResponse;
                            if (httpResponse != null && httpResponse.StatusCode == HttpStatusCode.Unauthorized)
                            {
                                // Ignore any security exception and simply keep 
                                // the AppCatalog URI null
                            }
                            else
                            {
                                throw ex;
                            }
                        }
                        else
                        {
                            throw ex;
                        }
                    }

                    if (appCatalogUri != null)
                    {
                        var manager = new AppManager(context);

                        foreach (var app in provisioningTenant.AppCatalog.Packages)
                        {

                            AppMetadata appMetadata = null;

                            if (app.Action == PackageAction.Upload || app.Action == PackageAction.UploadAndPublish)
                            {
                                var appSrc = parser.ParseString(app.Src);
                                var appBytes = ConnectorFileHelper.GetFileBytes(connector, appSrc);

                                var hash = string.Empty;
                                using (var memoryStream = new MemoryStream(appBytes))
                                {
                                    hash = CalculateHash(memoryStream);
                                }

                                var exists = false;
                                var appId = Guid.Empty;

                                using (var appCatalogContext = ((ClientContext)tenant.Context).Clone(appCatalogUri, applyingInformation.AccessTokens))
                                {
                                    // check if the app already is present
                                    var appList = appCatalogContext.Web.GetListByUrl("AppCatalog");
                                    var camlQuery = new CamlQuery
                                    {
                                        ViewXml = string.Format(appExistsQuery, hash)
                                    };
                                    var items = appList.GetItems(camlQuery);
                                    appCatalogContext.Load(items, i => i.IncludeWithDefaultProperties());
                                    appCatalogContext.ExecuteQueryRetry();
                                    if (items.Count > 0)
                                    {
                                        exists = true;
                                        appId = Guid.Parse(items[0].FieldValues["UniqueId"].ToString());
                                    }

                                }
                                var appFilename = appSrc.Substring(appSrc.LastIndexOf('\\') + 1);

                                if (!exists)
                                {
                                    messagesDelegate?.Invoke($"Processing solution {app.Src}", ProvisioningMessageType.Progress);
                                    appMetadata = manager.Add(appBytes, appFilename, app.Overwrite, timeoutSeconds: 500);
                                }
                                else
                                {
                                    messagesDelegate?.Invoke($"Skipping existing solution {app.Src}", ProvisioningMessageType.Progress);
                                    appMetadata = manager.GetAvailable().FirstOrDefault(a => a.Id == appId);
                                }
                                if (appMetadata != null)
                                {
                                    parser.AddToken(new AppPackageIdToken(web, appFilename, appMetadata.Id));
                                    parser.AddToken(new AppPackageIdToken(web, appMetadata.Title, appMetadata.Id));
                                }
                            }

                            if (app.Action == PackageAction.Publish || app.Action == PackageAction.UploadAndPublish)
                            {
                                if (appMetadata == null)
                                {
                                    appMetadata = manager.GetAvailable()
                                        .FirstOrDefault(a => a.Id == Guid.Parse(parser.ParseString(app.PackageId)));
                                }
                                if (appMetadata != null)
                                {
                                    manager.Deploy(appMetadata, app.SkipFeatureDeployment);
                                }
                                else
                                {
                                    scope.LogError("Referenced App Package {0} not available", app.PackageId);
                                    throw new Exception($"Referenced App Package {app.PackageId} not available");
                                }
                            }

                            if (app.Action == PackageAction.Remove)
                            {
                                var appId = Guid.Parse(parser.ParseString(app.PackageId));

                                // Get the apps already installed in the site
                                var appExists = manager.GetAvailable()?.Any(a => a.Id == appId);

                                if (appExists.HasValue && appExists.Value)
                                {
                                    manager.Remove(appId);
                                }
                                else
                                {
                                    messagesDelegate?.Invoke($"App Package with ID {appId} does not exist in the AppCatalog and cannot be removed!", ProvisioningMessageType.Warning);
                                }
                            }
                        }
                    }
                    else
                    {
                        messagesDelegate?.Invoke($"Tenant app catalog doesn't exist. ALM step will be skipped!", ProvisioningMessageType.Warning);
                    }
                }

            }
            return parser;
        }

        internal static string CalculateHash(Stream stream)
        {
            SHA512Managed HashCalculator = new SHA512Managed();
            byte[] packageHash = HashCalculator.ComputeHash(stream);
            string hashString = Convert.ToBase64String(packageHash);
            return hashString;
        }

        internal static TokenParser ProcessStorageEntities(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope, ProvisioningTemplateApplyingInformation applyingInformation, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.StorageEntities != null && provisioningTenant.StorageEntities.Any())
            {
                var rootSiteUrl = tenant.GetRootSiteUrl();
                tenant.Context.ExecuteQueryRetry();

                using (var context = ((ClientContext)tenant.Context).Clone(rootSiteUrl.Value, applyingInformation.AccessTokens))
                {
                    var web = context.Web;

                    Uri appCatalogUri = null;

                    try
                    {
                        appCatalogUri = web.GetAppCatalog();
                    }
                    catch (System.Net.WebException ex)
                    {
                        if (ex.Response != null)
                        {
                            var httpResponse = ex.Response as System.Net.HttpWebResponse;
                            if (httpResponse != null && httpResponse.StatusCode == HttpStatusCode.Unauthorized)
                            {
                                // Ignore any security exception and simply keep 
                                // the AppCatalog URI null
                            }
                            else
                            {
                                throw ex;
                            }
                        }
                        else
                        {
                            throw ex;
                        }
                    }

                    if (appCatalogUri != null)
                    {
                        using (var appCatalogContext = context.Clone(appCatalogUri, applyingInformation.AccessTokens))
                        {
                            foreach (var entity in provisioningTenant.StorageEntities)
                            {
                                var key = parser.ParseString(entity.Key);
                                var value = parser.ParseString(entity.Value);
                                var description = parser.ParseString(entity.Description);
                                var comment = parser.ParseString(entity.Comment);
                                appCatalogContext.Web.SetStorageEntity(key, value, description, comment);
                            }
                            appCatalogContext.Web.Update();
                            appCatalogContext.ExecuteQueryRetry();
                        }
                    }
                    else
                    {
                        messagesDelegate?.Invoke($"Tenant app catalog doesn't exist. Provisioning of storage entities will be skipped!", ProvisioningMessageType.Warning);
                    }

                }
            }
            return parser;
        }

        internal static TokenParser ProcessSiteScripts(Tenant tenant, ProvisioningTenant provisioningTenant, FileConnectorBase connector, TokenParser parser, PnPMonitoredScope scope, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.SiteScripts != null && provisioningTenant.SiteScripts.Any())
            {
                var existingScripts = tenant.GetSiteScripts();
                tenant.Context.Load(existingScripts);
                tenant.Context.ExecuteQueryRetry();

                foreach (var siteScript in provisioningTenant.SiteScripts)
                {

                    var parsedTitle = parser.ParseString(siteScript.Title);
                    var parsedDescription = parser.ParseString(siteScript.Description);
                    var parsedContent = parser.ParseString(System.Text.Encoding.UTF8.GetString(ConnectorFileHelper.GetFileBytes(connector, parser.ParseString(siteScript.JsonFilePath))));
                    var existingScript = existingScripts.FirstOrDefault(s => s.Title == parsedTitle);

                    messagesDelegate?.Invoke($"Processing site script {parsedTitle}", ProvisioningMessageType.Progress);

                    if (existingScript == null)
                    {
                        TenantSiteScriptCreationInfo siteScriptCreationInfo = new TenantSiteScriptCreationInfo
                        {
                            Title = parsedTitle,
                            Description = parsedDescription,
                            Content = parsedContent
                        };
                        var script = tenant.CreateSiteScript(siteScriptCreationInfo);
                        tenant.Context.Load(script);
                        tenant.Context.ExecuteQueryRetry();
                        parser.AddToken(new SiteScriptIdToken(null, parsedTitle, script.Id));
                    }
                    else
                    {
                        if (siteScript.Overwrite)
                        {
                            var existingId = existingScript.Id;
                            existingScript = Tenant.GetSiteScript(tenant.Context, existingId);
                            tenant.Context.ExecuteQueryRetry();

                            existingScript.Content = parsedContent;
                            existingScript.Title = parsedTitle;
                            existingScript.Description = parsedDescription;
                            tenant.UpdateSiteScript(existingScript);
                            tenant.Context.ExecuteQueryRetry();
                            var existingToken = parser.Tokens.OfType<SiteScriptIdToken>().FirstOrDefault(t => t.GetReplaceValue() == existingId.ToString());
                            if (existingToken != null)
                            {
                                parser.Tokens.Remove(existingToken);
                            }
                            parser.AddToken(new SiteScriptIdToken(null, parsedTitle, existingId));
                        }
                    }
                }
            }
            return parser;
        }

        public static TokenParser ProcessSiteDesigns(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.SiteDesigns != null && provisioningTenant.SiteDesigns.Any())
            {

                var existingDesigns = tenant.GetSiteDesigns();
                tenant.Context.Load(existingDesigns);
                tenant.Context.ExecuteQueryRetry();
                foreach (var siteDesign in provisioningTenant.SiteDesigns)
                {
                    var parsedTitle = parser.ParseString(siteDesign.Title);
                    var parsedDescription = parser.ParseString(siteDesign.Description);
                    var parsedPreviewImageUrl = parser.ParseString(siteDesign.PreviewImageUrl);
                    var parsedPreviewImageAltText = parser.ParseString(siteDesign.PreviewImageAltText);
                    messagesDelegate?.Invoke($"Processing site design {parsedTitle}", ProvisioningMessageType.Progress);

                    var existingSiteDesign = existingDesigns.FirstOrDefault(d => d.Title == parsedTitle);
                    if (existingSiteDesign == null)
                    {
                        TenantSiteDesignCreationInfo siteDesignCreationInfo = new TenantSiteDesignCreationInfo()
                        {
                            Title = parsedTitle,
                            Description = parsedDescription,
                            PreviewImageUrl = parsedPreviewImageUrl,
                            PreviewImageAltText = parsedPreviewImageAltText,
                            IsDefault = siteDesign.IsDefault,
                        };
                        switch ((int)siteDesign.WebTemplate)
                        {
                            case 0:
                                {
                                    siteDesignCreationInfo.WebTemplate = "64";
                                    break;
                                }
                            case 1:
                                {
                                    siteDesignCreationInfo.WebTemplate = "68";
                                    break;
                                }
                        }
                        if (siteDesign.SiteScripts != null && siteDesign.SiteScripts.Any())
                        {
                            List<Guid> ids = new List<Guid>();
                            foreach (var siteScriptRef in siteDesign.SiteScripts)
                            {
                                ids.Add(Guid.Parse(parser.ParseString(siteScriptRef)));
                            }
                            siteDesignCreationInfo.SiteScriptIds = ids.ToArray();
                        }
                        var design = tenant.CreateSiteDesign(siteDesignCreationInfo);
                        tenant.Context.Load(design);
                        tenant.Context.ExecuteQueryRetry();

                        if (siteDesign.Grants != null && siteDesign.Grants.Any())
                        {
                            foreach (var grant in siteDesign.Grants)
                            {
                                var rights = (TenantSiteDesignPrincipalRights)Enum.Parse(typeof(TenantSiteDesignPrincipalRights), grant.Right.ToString());
                                tenant.GrantSiteDesignRights(design.Id, new[] { grant.Principal }, rights);
                            }
                            tenant.Context.ExecuteQueryRetry();
                        }
                        parser.AddToken(new SiteDesignIdToken(null, design.Title, design.Id));
                    }
                    else
                    {
                        if (siteDesign.Overwrite)
                        {
                            var existingId = existingSiteDesign.Id;
                            existingSiteDesign = Tenant.GetSiteDesign(tenant.Context, existingId);
                            tenant.Context.ExecuteQueryRetry();

                            existingSiteDesign.Title = parsedTitle;
                            existingSiteDesign.Description = parsedDescription;
                            existingSiteDesign.PreviewImageUrl = parsedPreviewImageUrl;
                            existingSiteDesign.PreviewImageAltText = parsedPreviewImageAltText;
                            existingSiteDesign.IsDefault = siteDesign.IsDefault;
                            switch ((int)siteDesign.WebTemplate)
                            {
                                case 0:
                                    {
                                        existingSiteDesign.WebTemplate = "64";
                                        break;
                                    }
                                case 1:
                                    {
                                        existingSiteDesign.WebTemplate = "68";
                                        break;
                                    }
                            }

                            tenant.UpdateSiteDesign(existingSiteDesign);
                            tenant.Context.ExecuteQueryRetry();

                            var existingToken = parser.Tokens.OfType<SiteDesignIdToken>().FirstOrDefault(t => t.GetReplaceValue() == existingId.ToString());
                            if (existingToken != null)
                            {
                                parser.Tokens.Remove(existingToken);
                            }
                            parser.AddToken(new SiteScriptIdToken(null, parsedTitle, existingId));

                            if (siteDesign.Grants != null && siteDesign.Grants.Any())
                            {
                                var existingRights = Tenant.GetSiteDesignRights(tenant.Context, existingId);
                                tenant.Context.Load(existingRights);
                                tenant.Context.ExecuteQueryRetry();
                                foreach (var existingRight in existingRights)
                                {
                                    Tenant.RevokeSiteDesignRights(tenant.Context, existingId, new[] { existingRight.PrincipalName });
                                }
                                foreach (var grant in siteDesign.Grants)
                                {
                                    var rights = (TenantSiteDesignPrincipalRights)Enum.Parse(typeof(TenantSiteDesignPrincipalRights), grant.Right.ToString());
                                    tenant.GrantSiteDesignRights(existingId, new[] { parser.ParseString(grant.Principal) }, rights);
                                }
                                tenant.Context.ExecuteQueryRetry();
                            }
                        }
                    }
                }
            }
            return parser;
        }

        public static TokenParser ProcessThemes(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.Themes != null && provisioningTenant.Themes.Any())
            {
                foreach (var theme in provisioningTenant.Themes)
                {
                    var parsedName = parser.ParseString(theme.Name);
                    var parsedPalette = parser.ParseString(theme.Palette);

                    messagesDelegate?.Invoke($"Processing theme {parsedName}", ProvisioningMessageType.Progress);

                    var palette = JsonConvert.DeserializeObject<Dictionary<string, string>>(parsedPalette);
                    var tenantTheme = new TenantTheme() { Name = parsedName, Palette = palette, IsInverted = theme.IsInverted };
                    tenant.UpdateTenantTheme(parsedName, JsonConvert.SerializeObject(tenantTheme));
                    tenant.Context.ExecuteQueryRetry();
                }
            }
            return parser;
        }

        public static TokenParser ProcessWebApiPermissions(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.WebApiPermissions != null && provisioningTenant.WebApiPermissions.Any())
            {
                messagesDelegate?.Invoke("Processing WebApiPermissions", ProvisioningMessageType.Progress);
                var servicePrincipal = new SPOWebAppServicePrincipal(tenant.Context);
                //var requests = servicePrincipal.PermissionRequests;
                var requestsEnumerable = tenant.Context.LoadQuery(servicePrincipal.PermissionRequests);
                var grantsEnumerable = tenant.Context.LoadQuery(servicePrincipal.PermissionGrants);
                tenant.Context.ExecuteQueryRetry();

                var requests = requestsEnumerable.ToList();

                foreach (var permission in provisioningTenant.WebApiPermissions)
                {
                    var parsedScope = parser.ParseString(permission.Scope);
                    var parsedResource = parser.ParseString(permission.Resource);
                    var request = requests.FirstOrDefault(r => r.Scope.Equals(parsedScope, StringComparison.InvariantCultureIgnoreCase) && r.Resource.Equals(parsedResource, StringComparison.InvariantCultureIgnoreCase));
                    while (request != null)
                    {
                        if (grantsEnumerable.FirstOrDefault(g => g.Resource.Equals(parsedResource, StringComparison.InvariantCultureIgnoreCase) && g.Scope.ToLower().Contains(parsedScope.ToLower())) == null)
                        {
                            var requestToApprove = servicePrincipal.PermissionRequests.GetById(request.Id);
                            tenant.Context.Load(requestToApprove);
                            tenant.Context.ExecuteQueryRetry();
                            try
                            {
                                requestToApprove.Approve();
                                tenant.Context.ExecuteQueryRetry();
                            }
                            catch (Exception ex)
                            {
                                messagesDelegate?.Invoke(ex.Message, ProvisioningMessageType.Warning);
                            }
                        }
                        requests.Remove(request);
                        request = requests.FirstOrDefault(r => r.Scope.Equals(parsedScope, StringComparison.InvariantCultureIgnoreCase) && r.Resource.Equals(parsedResource, StringComparison.InvariantCultureIgnoreCase));
                    }
                }
            }
            return parser;
        }

        [DataContract]
        internal class TenantTheme
        {
            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "palette")]
            public IDictionary<string, string> Palette { get; set; }

            [DataMember(Name = "isInverted")]
            public bool IsInverted { get; set; }
        }
        /// <summary>
        /// Retrieves a file as a byte array from the connector. If the file name contains special characters (e.g. "%20") and cannot be retrieved, a workaround will be performed
        /// </summary>


        internal static void ProcessCdns(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.ContentDeliveryNetwork != null)
            {
                if (provisioningTenant.ContentDeliveryNetwork.PublicCdn != null || provisioningTenant.ContentDeliveryNetwork.PrivateCdn != null)
                {
                    var publicCdnEnabled = tenant.GetTenantCdnEnabled(SPOTenantCdnType.Public);
                    var privateCdnEnabled = tenant.GetTenantCdnEnabled(SPOTenantCdnType.Private);
                    tenant.Context.ExecuteQueryRetry();
                    var publicCdn = provisioningTenant.ContentDeliveryNetwork.PublicCdn;
                    if (publicCdn != null)
                    {
                        if (publicCdnEnabled.Value != publicCdn.Enabled)
                        {
                            scope.LogInfo($"Public CDN is set to {(publicCdn.Enabled ? "Enabled" : "Disabled")}");
                            tenant.SetTenantCdnEnabled(SPOTenantCdnType.Public, publicCdn.Enabled);
                            tenant.Context.ExecuteQueryRetry();
                        }
                        if (publicCdn.Enabled)
                        {
                            if (!publicCdn.NoDefaultOrigins)
                            {
                                tenant.CreateTenantCdnDefaultOrigins(SPOTenantCdnType.Public);
                                tenant.Context.ExecuteQueryRetry();
                            }
                            ProcessOrigins(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
                            ProcessPolicies(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
                        }
                    }
                    var privateCdn = provisioningTenant.ContentDeliveryNetwork.PrivateCdn;
                    if (privateCdn != null)
                    {
                        if (privateCdnEnabled.Value != privateCdn.Enabled)
                        {
                            scope.LogInfo($"Private CDN is set to {(privateCdn.Enabled ? "Enabled" : "Disabled")}");
                            tenant.SetTenantCdnEnabled(SPOTenantCdnType.Private, privateCdn.Enabled);
                            tenant.Context.ExecuteQueryRetry();
                        }
                        if (privateCdn.Enabled)
                        {
                            if (!privateCdn.NoDefaultOrigins)
                            {
                                tenant.CreateTenantCdnDefaultOrigins(SPOTenantCdnType.Private);
                                tenant.Context.ExecuteQueryRetry();
                            }
                            ProcessOrigins(tenant, privateCdn, SPOTenantCdnType.Private, parser, scope);
                            ProcessPolicies(tenant, privateCdn, SPOTenantCdnType.Private, parser, scope);
                        }
                    }
                }
            }
        }

        public static void ProcessOrigins(Tenant tenant, CdnSettings cdnSettings, SPOTenantCdnType cdnType, TokenParser parser, PnPMonitoredScope scope)
        {
            if (cdnSettings.Origins != null && cdnSettings.Origins.Any())
            {
                var origins = tenant.GetTenantCdnOrigins(cdnType);
                tenant.Context.ExecuteQueryRetry();
                foreach (var origin in cdnSettings.Origins)
                {
                    switch (origin.Action)
                    {
                        case OriginAction.Add:
                            {
                                var parsedOriginUrl = parser.ParseString(origin.Url);
                                if (!origins.Contains(parsedOriginUrl))
                                {
                                    scope.LogInfo($"Adding {parsedOriginUrl} to {cdnType} CDN");
                                    tenant.AddTenantCdnOrigin(cdnType, parsedOriginUrl);
                                }
                                break;
                            }
                        case OriginAction.Remove:
                            {
                                var parsedOriginUrl = parser.ParseString(origin.Url);
                                if (origins.Contains(parsedOriginUrl))
                                {
                                    scope.LogInfo($"Removing {parsedOriginUrl} to {cdnType} CDN");
                                    tenant.RemoveTenantCdnOrigin(cdnType, parsedOriginUrl);
                                }
                                break;
                            }
                    }
                    tenant.Context.ExecuteQueryRetry();
                }
            }
        }

        public static void ProcessPolicies(Tenant tenant, CdnSettings cdnSettings, SPOTenantCdnType cdnType, TokenParser parser, PnPMonitoredScope scope)
        {
            var isDirty = false;
            var rawPolicies = tenant.GetTenantCdnPolicies(cdnType);
            tenant.Context.ExecuteQueryRetry();
            var policies = ParsePolicies(rawPolicies);

            if (!string.IsNullOrEmpty(cdnSettings.IncludeFileExtensions))
            {

                var parsedValue = parser.ParseString(cdnSettings.IncludeFileExtensions);
                if (policies.FirstOrDefault(p => p.Key == SPOTenantCdnPolicyType.IncludeFileExtensions).Value != parsedValue)
                {
                    scope.LogInfo($"Setting IncludeFileExtensions policy to {parsedValue}");
                    tenant.SetTenantCdnPolicy(cdnType, SPOTenantCdnPolicyType.IncludeFileExtensions, parsedValue);
                    isDirty = true;
                }
            }
            if (!string.IsNullOrEmpty(cdnSettings.ExcludeRestrictedSiteClassifications))
            {
                var parsedValue = parser.ParseString(cdnSettings.ExcludeRestrictedSiteClassifications);
                if (policies.FirstOrDefault(p => p.Key == SPOTenantCdnPolicyType.ExcludeRestrictedSiteClassifications).Value != parsedValue)
                {
                    scope.LogInfo($"Setting ExcludeRestrictSiteClassifications policy to {parsedValue}");
                    tenant.SetTenantCdnPolicy(cdnType, SPOTenantCdnPolicyType.ExcludeRestrictedSiteClassifications, parsedValue);
                    isDirty = true;
                }
            }
            if (!string.IsNullOrEmpty(cdnSettings.ExcludeIfNoScriptDisabled))
            {

                var parsedValue = parser.ParseString(cdnSettings.ExcludeIfNoScriptDisabled);
                if (policies.FirstOrDefault(p => p.Key == SPOTenantCdnPolicyType.ExcludeIfNoScriptDisabled).Value != parsedValue)
                {
                    scope.LogInfo($"Setting ExcludeIfNoScriptDisabled policy to {parsedValue}");
                    tenant.SetTenantCdnPolicy(cdnType, SPOTenantCdnPolicyType.ExcludeIfNoScriptDisabled, parsedValue);
                    isDirty = true;
                }
            }
            if (isDirty)
            {
                tenant.Context.ExecuteQueryRetry();
            }
        }

        private static Dictionary<Microsoft.Online.SharePoint.TenantAdministration.SPOTenantCdnPolicyType, string> ParsePolicies(IList<string> entries)
        {
            var returnDict = new Dictionary<SPOTenantCdnPolicyType, string>();
            foreach (var entry in entries)
            {
                var entryArray = entry.Split(new[] { ';' });
                returnDict.Add((SPOTenantCdnPolicyType)Enum.Parse(typeof(SPOTenantCdnPolicyType), entryArray[0]), entryArray[1]);
            }
            return returnDict;
        }
    }
}
#endif