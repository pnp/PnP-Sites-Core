#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
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

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    internal static class TenantHelper
    {
        public static TokenParser ProcessApps(Tenant tenant, ProvisioningTenant provisioningTenant, FileConnectorBase connector, TokenParser parser, PnPMonitoredScope scope, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (provisioningTenant.AppCatalog != null && provisioningTenant.AppCatalog.Packages.Count > 0)
            {
                var rootSiteUrl = tenant.GetRootSiteUrl();
                tenant.Context.ExecuteQueryRetry();
                using (var context = ((ClientContext)tenant.Context).Clone(rootSiteUrl.Value))
                {
                    var web = context.Web;
                    var appCatalogUri = web.GetAppCatalog();

                    var manager = new AppManager(context);

                    if (appCatalogUri != null)
                    {
                        foreach (var app in provisioningTenant.AppCatalog.Packages)
                        {
                            AppMetadata appMetadata = null;

                            if (app.Action == PackageAction.Upload || app.Action == PackageAction.UploadAndPublish)
                            {
                                var appSrc = parser.ParseString(app.Src);
                                var appBytes = GetFileBytes(connector, appSrc);

                                var appFilename = appSrc.Substring(appSrc.LastIndexOf('\\') + 1);
                                appMetadata = manager.Add(appBytes, appFilename, app.Overwrite, timeoutSeconds: 300);

                                parser.Tokens.Add(new AppPackageIdToken(web, appFilename, appMetadata.Id));
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

        internal static TokenParser ProcessStorageEntities(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope)
        {
            if (provisioningTenant.StorageEntities != null && provisioningTenant.StorageEntities.Any())
            {
                using (var context = ((ClientContext)tenant.Context).Clone(tenant.RootSiteUrl))
                {
                    var web = context.Web;
                    var appCatalogUri = web.GetAppCatalog();
                    using (var appCatalogContext = context.Clone(appCatalogUri))
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
            }
            return parser;
        }

        internal static TokenParser ProcessSiteScripts(Tenant tenant, ProvisioningTenant provisioningTenant, FileConnectorBase connector, TokenParser parser, PnPMonitoredScope scope)
        {
            if (provisioningTenant.SiteScripts != null && provisioningTenant.SiteScripts.Any())
            {
                var existingScripts = tenant.GetSiteScripts();
                tenant.Context.Load(existingScripts);
                tenant.Context.ExecuteQueryRetry();

                foreach (var siteScript in provisioningTenant.SiteScripts)
                {
                    var scriptTitle = parser.ParseString(siteScript.Title);
                    var scriptDescription = parser.ParseString(siteScript.Description);
                    var scriptContent = parser.ParseString(System.Text.Encoding.UTF8.GetString(GetFileBytes(connector, parser.ParseString(siteScript.JsonFilePath))));
                    var existingScript = existingScripts.FirstOrDefault(s => s.Title == scriptTitle);

                    if (existingScript == null)
                    {
                        TenantSiteScriptCreationInfo siteScriptCreationInfo = new TenantSiteScriptCreationInfo
                        {
                            Title = scriptTitle,
                            Description = scriptDescription,
                            Content = scriptContent
                        };
                        var script = tenant.CreateSiteScript(siteScriptCreationInfo);
                        tenant.Context.Load(script);
                        tenant.Context.ExecuteQueryRetry();
                        parser.AddToken(new SiteScriptIdToken(null, scriptTitle, script.Id));
                    }
                    else
                    {
                        if (siteScript.Overwrite)
                        {
                            var existingId = existingScript.Id;
                            existingScript = Tenant.GetSiteScript(tenant.Context, existingId);
                            tenant.Context.ExecuteQueryRetry();

                            existingScript.Content = scriptContent;
                            existingScript.Title = scriptTitle;
                            existingScript.Description = scriptDescription;
                            tenant.UpdateSiteScript(existingScript);
                            tenant.Context.ExecuteQueryRetry();
                            var existingToken = parser.Tokens.OfType<SiteScriptIdToken>().FirstOrDefault(t => t.GetReplaceValue() == existingId.ToString());
                            if (existingToken != null)
                            {
                                parser.Tokens.Remove(existingToken);
                            }
                            parser.AddToken(new SiteScriptIdToken(null, scriptTitle, existingId));
                        }
                    }
                }
            }
            return parser;
        }

        public static TokenParser ProcessSiteDesigns(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope)
        {
            if (provisioningTenant.SiteDesigns != null && provisioningTenant.SiteDesigns.Any())
            {

                var existingDesigns = tenant.GetSiteDesigns();
                tenant.Context.Load(existingDesigns);
                tenant.Context.ExecuteQueryRetry();
                foreach (var siteDesign in provisioningTenant.SiteDesigns)
                {
                    var designTitle = parser.ParseString(siteDesign.Title);
                    var designDescription = parser.ParseString(siteDesign.Description);
                    var designPreviewImageUrl = parser.ParseString(siteDesign.PreviewImageUrl);
                    var designPreviewImageAltText = parser.ParseString(siteDesign.PreviewImageAltText);

                    var existingSiteDesign = existingDesigns.FirstOrDefault(d => d.Title == designTitle);
                    if (existingSiteDesign == null)
                    {
                        TenantSiteDesignCreationInfo siteDesignCreationInfo = new TenantSiteDesignCreationInfo()
                        {
                            Title = designTitle,
                            Description = designDescription,
                            PreviewImageUrl = designPreviewImageUrl,
                            PreviewImageAltText = designPreviewImageAltText,
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

                            existingSiteDesign.Title = designTitle;
                            existingSiteDesign.Description = designDescription;
                            existingSiteDesign.PreviewImageUrl = designPreviewImageUrl;
                            existingSiteDesign.PreviewImageAltText = designPreviewImageAltText;
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
                            parser.AddToken(new SiteScriptIdToken(null, designTitle, existingId));

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

        public static TokenParser ProcessThemes(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope)
        {
            if (provisioningTenant.Themes != null && provisioningTenant.Themes.Any())
            {
                foreach (var theme in provisioningTenant.Themes)
                {
                    var parsedName = parser.ParseString(theme.Name);
                    var parsedPalette = parser.ParseString(theme.Palette);
                    var palette = JsonConvert.DeserializeObject<Dictionary<string, string>>(parsedPalette);
                    var tenantTheme = new TenantTheme() { Name = parsedName, Palette = palette, IsInverted = theme.IsInverted };
                    tenant.UpdateTenantTheme(parsedName, JsonConvert.SerializeObject(tenantTheme));
                    tenant.Context.ExecuteQueryRetry();
                }
            }
            return parser;
        }


        [DataContract]
        private class TenantTheme
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
        private static byte[] GetFileBytes(FileConnectorBase connector, string fileName)
        {
            var container = String.Empty;
            if (fileName.Contains(@"\") || fileName.Contains(@"/"))
            {
                var tempFileName = fileName.Replace(@"/", @"\");
                container = fileName.Substring(0, tempFileName.LastIndexOf(@"\"));
                fileName = fileName.Substring(tempFileName.LastIndexOf(@"\") + 1);
            }

            // add the default provided container (if any)
            if (!String.IsNullOrEmpty(container))
            {
                if (!String.IsNullOrEmpty(connector.GetContainer()))
                {
                    if (container.StartsWith("/"))
                    {
                        container = container.TrimStart("/".ToCharArray());
                    }

#if !NETSTANDARD2_0
                    if (connector.GetType() == typeof(Connectors.AzureStorageConnector))
                    {
                        if (connector.GetContainer().EndsWith("/"))
                        {
                            container = $@"{connector.GetContainer()}{container}";
                        }
                        else
                        {
                            container = $@"{connector.GetContainer()}/{container}";
                        }
                    }
                    else
                    {
                        container = $@"{connector.GetContainer()}\{container}";
                    }
#else
                    container = $@"{template.Connector.GetContainer()}\{container}";
#endif
                }
            }
            else
            {
                container = connector.GetContainer();
            }

            var stream = connector.GetFileStream(fileName, container);
            if (stream == null)
            {
                //Decode the URL and try again
                fileName = WebUtility.UrlDecode(fileName);
                stream = connector.GetFileStream(fileName, container);
            }
            byte[] returnData;

            using (var memStream = new MemoryStream())
            {
                stream.CopyTo(memStream);
                memStream.Position = 0;
                returnData = memStream.ToArray();
            }
            if (stream != null)
            {
                stream.Dispose();
            }
            return returnData;
        }

        internal static void ProcessCdns(Tenant tenant, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope)
        {
            if (provisioningTenant.ContentDeliveryNetwork != null)
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
                        ProcessOrigins(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
                        ProcessPolicies(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
                    }
                }
                var privateCdn = provisioningTenant.ContentDeliveryNetwork.PrivateCdn;
                if (privateCdn != null)
                {
                    if (privateCdnEnabled.Value != privateCdn.Enabled)
                    {
                        scope.LogInfo($"Private CDN is set to {(publicCdn.Enabled ? "Enabled" : "Disabled")}");
                        tenant.SetTenantCdnEnabled(SPOTenantCdnType.Private, privateCdn.Enabled);
                        tenant.Context.ExecuteQueryRetry();
                    }
                    if (privateCdn.Enabled)
                    {
                        ProcessOrigins(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
                        ProcessPolicies(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
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