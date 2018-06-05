using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.ALM;
using System.IO;
using System.Net;
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
#if !ONPREMISES
    internal class ObjectTenant : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Tenant Settings"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            web.EnsureProperty(w => w.Url);

            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Tenant != null)
                {
                    ProcessCdns(web, template.Tenant, parser, scope);
                    ProcessApps(web, template, parser, scope);
                    parser = ProcessSiteScripts(web, template, parser, scope);
                    parser = ProcessSiteDesigns(web, template, parser, scope);
                    parser = ProcessStorageEntities(web, template, parser, scope);
                    // So far we do not provision CDN settings
                    // It will come in the near future
                    // NOOP on CDN
                }
            }

            return parser;
        }

        private void ProcessApps(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope)
        {
            if (template.Tenant.AppCatalog != null && template.Tenant.AppCatalog.Packages.Count > 0)
            {
                var manager = new AppManager(web.Context as ClientContext);

                var appCatalogUri = web.GetAppCatalog();
                if (appCatalogUri != null)
                {
                    foreach (var app in template.Tenant.AppCatalog.Packages)
                    {
                        AppMetadata appMetadata = null;

                        if (app.Action == PackageAction.Upload || app.Action == PackageAction.UploadAndPublish)
                        {
                            var appBytes = GetFileBytes(template, app.Src);

                            var appFilename = app.Src.Substring(app.Src.LastIndexOf('\\') + 1);
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
                                WriteMessage($"App Package with ID {appId} does not exist in the AppCatalog and cannot be removed!", ProvisioningMessageType.Warning);
                            }
                        }
                    }
                }
                else
                {
                    WriteMessage($"Tenant app catalog doesn't exist. ALM step will be skipped!", ProvisioningMessageType.Warning);
                }
            }
        }

        private TokenParser ProcessStorageEntities(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope)
        {
            if (template.Tenant.StorageEntities != null && template.Tenant.StorageEntities.Any())
            {
                var appCatalogUri = web.GetAppCatalog();
                using (var appCatalogContext = web.Context.Clone(appCatalogUri))
                {
                    foreach (var entity in template.Tenant.StorageEntities)
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
            return parser;
        }

        private TokenParser ProcessSiteScripts(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope)
        {
            if (template.Tenant.SiteScripts != null && template.Tenant.SiteScripts.Any())
            {
                using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl()))
                {
                    var tenant = new Tenant(tenantContext);

                    var existingScripts = tenant.GetSiteScripts();
                    tenantContext.Load(existingScripts);
                    tenantContext.ExecuteQueryRetry();

                    foreach (var siteScript in template.Tenant.SiteScripts)
                    {
                        var scriptTitle = parser.ParseString(siteScript.Title);
                        var scriptDescription = parser.ParseString(siteScript.Description);
                        var scriptContent = parser.ParseString(System.Text.Encoding.UTF8.GetString(GetFileBytes(template, parser.ParseString(siteScript.JsonFilePath))));
                        var existingScript = existingScripts.FirstOrDefault(s => s.Title == scriptTitle);

                        if (existingScript == null)
                        {
                            TenantSiteScriptCreationInfo siteScriptCreationInfo = new TenantSiteScriptCreationInfo
                            {
                                Title = siteScript.Title,
                                Description = siteScript.Description,
                                Content = scriptContent
                            };
                            var script = tenant.CreateSiteScript(siteScriptCreationInfo);
                            tenantContext.Load(script);
                            tenantContext.ExecuteQueryRetry();
                            parser.AddToken(new SiteScriptIdToken(web, scriptTitle, script.Id));
                        }
                        else
                        {
                            if (siteScript.Overwrite)
                            {
                                var existingId = existingScript.Id;
                                existingScript = Tenant.GetSiteScript(tenantContext, existingId);
                                tenantContext.ExecuteQueryRetry();

                                existingScript.Content = scriptContent;
                                existingScript.Title = scriptTitle;
                                existingScript.Description = scriptDescription;
                                tenant.UpdateSiteScript(existingScript);
                                tenantContext.ExecuteQueryRetry();
                                var existingToken = parser.Tokens.OfType<SiteScriptIdToken>().FirstOrDefault(t => t.GetReplaceValue() == existingId.ToString());
                                if (existingToken != null)
                                {
                                    parser.Tokens.Remove(existingToken);
                                }
                                parser.AddToken(new SiteScriptIdToken(web, scriptTitle, existingId));
                            }
                        }
                    }
                }
            }
            return parser;
        }

        private TokenParser ProcessSiteDesigns(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope)
        {
            if (template.Tenant.SiteDesigns != null && template.Tenant.SiteDesigns.Any())
            {
                using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl()))
                {
                    var tenant = new Tenant(tenantContext);

                    var existingDesigns = tenant.GetSiteDesigns();
                    tenantContext.Load(existingDesigns);
                    tenantContext.ExecuteQueryRetry();
                    foreach (var siteDesign in template.Tenant.SiteDesigns)
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
                            switch(siteDesign.WebTemplate)
                            {
                                case SiteDesignWebTemplate.TeamSite:
                                    {
                                        siteDesignCreationInfo.WebTemplate = "64";
                                        break;
                                    }
                                case SiteDesignWebTemplate.CommunicationSite:
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
                            tenantContext.Load(design);
                            tenantContext.ExecuteQueryRetry();

                            if (siteDesign.Grants != null && siteDesign.Grants.Any())
                            {
                                foreach (var grant in siteDesign.Grants)
                                {
                                    var rights = (TenantSiteDesignPrincipalRights)Enum.Parse(typeof(TenantSiteDesignPrincipalRights), grant.Right.ToString());
                                    tenant.GrantSiteDesignRights(design.Id, new[] { grant.Principal }, rights);
                                }
                                tenantContext.ExecuteQueryRetry();
                            }
                            parser.AddToken(new SiteDesignIdToken(web, design.Title, design.Id));
                        }
                        else
                        {
                            if (siteDesign.Overwrite)
                            {
                                var existingId = existingSiteDesign.Id;
                                existingSiteDesign = Tenant.GetSiteDesign(tenantContext, existingId);
                                tenantContext.ExecuteQueryRetry();

                                existingSiteDesign.Title = designTitle;
                                existingSiteDesign.Description = designDescription;
                                existingSiteDesign.PreviewImageUrl = designPreviewImageUrl;
                                existingSiteDesign.PreviewImageAltText = designPreviewImageAltText;
                                existingSiteDesign.IsDefault = siteDesign.IsDefault;
                                switch (siteDesign.WebTemplate)
                                {
                                    case SiteDesignWebTemplate.TeamSite:
                                        {
                                            existingSiteDesign.WebTemplate = "64";
                                            break;
                                        }
                                    case SiteDesignWebTemplate.CommunicationSite:
                                        {
                                            existingSiteDesign.WebTemplate = "68";
                                            break;
                                        }
                                }

                                tenant.UpdateSiteDesign(existingSiteDesign);
                                tenantContext.ExecuteQueryRetry();

                                var existingToken = parser.Tokens.OfType<SiteDesignIdToken>().FirstOrDefault(t => t.GetReplaceValue() == existingId.ToString());
                                if (existingToken != null)
                                {
                                    parser.Tokens.Remove(existingToken);
                                }
                                parser.AddToken(new SiteScriptIdToken(web, designTitle, existingId));

                                if (siteDesign.Grants != null && siteDesign.Grants.Any())
                                {
                                    var existingRights = Tenant.GetSiteDesignRights(tenantContext, existingId);
                                    tenantContext.Load(existingRights);
                                    tenantContext.ExecuteQueryRetry();
                                    foreach (var existingRight in existingRights)
                                    {
                                        Tenant.RevokeSiteDesignRights(tenantContext, existingId, new[] { existingRight.PrincipalName });
                                    }
                                    foreach (var grant in siteDesign.Grants)
                                    {
                                        var rights = (TenantSiteDesignPrincipalRights)Enum.Parse(typeof(TenantSiteDesignPrincipalRights), grant.Right.ToString());
                                        tenant.GrantSiteDesignRights(existingId, new[] { parser.ParseString(grant.Principal) }, rights);
                                    }
                                    tenantContext.ExecuteQueryRetry();
                                }
                            }
                        }
                    }
                }
            }
            return parser;
        }

        /// <summary>
        /// Retrieves a file as a byte array from the connector. If the file name contains special characters (e.g. "%20") and cannot be retrieved, a workaround will be performed
        /// </summary>
        private static byte[] GetFileBytes(ProvisioningTemplate template, string fileName)
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
                if (!String.IsNullOrEmpty(template.Connector.GetContainer()))
                {
                    if (container.StartsWith("/"))
                    {
                        container = container.TrimStart("/".ToCharArray());
                    }

#if !NETSTANDARD2_0
                    if (template.Connector.GetType() == typeof(Connectors.AzureStorageConnector))
                    {
                        if (template.Connector.GetContainer().EndsWith("/"))
                        {
                            container = $@"{template.Connector.GetContainer()}{container}";
                        }
                        else
                        {
                            container = $@"{template.Connector.GetContainer()}/{container}";
                        }
                    }
                    else
                    {
                        container = $@"{template.Connector.GetContainer()}\{container}";
                    }
#else
                    container = $@"{template.Connector.GetContainer()}\{container}";
#endif
                }
            }
            else
            {
                container = template.Connector.GetContainer();
            }

            var stream = template.Connector.GetFileStream(fileName, container);
            if (stream == null)
            {
                //Decode the URL and try again
                fileName = WebUtility.UrlDecode(fileName);
                stream = template.Connector.GetFileStream(fileName, container);
            }
            byte[] returnData;

            using (var memStream = new MemoryStream())
            {
                stream.CopyTo(memStream);
                memStream.Position = 0;
                returnData = memStream.ToArray();
            }
            return returnData;
        }

        private static void ProcessCdns(Web web, ProvisioningTenant provisioningTenant, TokenParser parser, PnPMonitoredScope scope)
        {
            if (provisioningTenant.ContentDeliveryNetwork != null)
            {
                using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl()))
                {
                    var tenant = new Tenant(tenantContext);
                    var publicCdnEnabled = tenant.GetTenantCdnEnabled(SPOTenantCdnType.Public);
                    var privateCdnEnabled = tenant.GetTenantCdnEnabled(SPOTenantCdnType.Private);
                    tenantContext.ExecuteQueryRetry();
                    var publicCdn = provisioningTenant.ContentDeliveryNetwork.PublicCdn;
                    if (publicCdn != null)
                    {
                        if (publicCdnEnabled.Value != publicCdn.Enabled)
                        {
                            scope.LogInfo($"Public CDN is set to {(publicCdn.Enabled ? "Enabled" : "Disabled")}");
                            tenant.SetTenantCdnEnabled(SPOTenantCdnType.Public, publicCdn.Enabled);
                            tenantContext.ExecuteQueryRetry();
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
                            tenantContext.ExecuteQueryRetry();
                        }
                        if (privateCdn.Enabled)
                        {
                            ProcessOrigins(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
                            ProcessPolicies(tenant, publicCdn, SPOTenantCdnType.Public, parser, scope);
                        }
                    }
                }
            }
        }

        private static void ProcessOrigins(Tenant tenant, CdnSettings cdnSettings, SPOTenantCdnType cdnType, TokenParser parser, PnPMonitoredScope scope)
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

        private static void ProcessPolicies(Tenant tenant, CdnSettings cdnSettings, SPOTenantCdnType cdnType, TokenParser parser, PnPMonitoredScope scope)
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

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            // By default we don't extract the packages
            return (false);
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return (template.Tenant != null);
        }
    }

#endif
}
