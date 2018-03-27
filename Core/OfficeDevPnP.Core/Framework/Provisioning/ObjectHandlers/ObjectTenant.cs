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
                    ProcessCdns(web, template.Tenant, parser);

                    var manager = new AppManager(web.Context as ClientContext);

                    if (template.Tenant.AppCatalog != null && template.Tenant.AppCatalog.Packages.Count > 0)
                    {
                        var appCatalogUri = web.GetAppCatalog();
                        if (appCatalogUri != null)
                        {
                            foreach (var app in template.Tenant.AppCatalog.Packages)
                            {
                                AppMetadata appMetadata = null;

                                if (app.Action == PackageAction.Upload ||
                                    app.Action == PackageAction.UploadAndPublish)
                                {
                                    using (var packageStream = GetPackageStream(template, app))
                                    {
                                        var memStream = new MemoryStream();
                                        packageStream.CopyTo(memStream);
                                        memStream.Position = 0;

                                        var appFilename = app.Src.Substring(app.Src.LastIndexOf('\\') + 1);
                                        appMetadata = manager.Add(memStream.ToArray(),
                                            appFilename,
                                            app.Overwrite);

                                        parser.Tokens.Add(new AppPackageIdToken(web, appFilename, appMetadata.Id));
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

                    // So far we do not provision CDN settings
                    // It will come in the near future
                    // NOOP on CDN
                }
            }

            return parser;
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

        /// <summary>
        /// Retrieves <see cref="Stream"/> from connector. If the file name contains special characters (e.g. "%20") and cannot be retrieved, a workaround will be performed
        /// </summary>
        private static Stream GetPackageStream(ProvisioningTemplate template, Model.Package package)
        {
            var fileName = package.Src;
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

            return stream;
        }

        private static void ProcessCdns(Web web, ProvisioningTenant provisioningTenant, TokenParser parser)
        {
            if (provisioningTenant.ContentDeliveryNetwork != null)
            {
                var webUrl = web.Url;
                var uri = new Uri(webUrl);
                if (!uri.Host.Contains("-admin."))
                {
                    var splittedHost = uri.Host.Split(new[] { '.' });
                    webUrl = $"{uri.Scheme}://{splittedHost[0]}-admin.{string.Join(".", splittedHost.Skip(1))}";
                }

                using (var tenantContext = web.Context.Clone(webUrl))
                {
                    var tenant = new Tenant(tenantContext);
                    var publicCdnEnabled = tenant.GetTenantCdnEnabled(SPOTenantCdnType.Public);
                    var privateCdnEnabled = tenant.GetTenantCdnEnabled(SPOTenantCdnType.Private);
                    tenantContext.ExecuteQueryRetry();
                    var publicCdn = provisioningTenant.ContentDeliveryNetwork.PublicCdn;
                    if (publicCdn != null)
                    {
                        tenant.SetTenantCdnEnabled(SPOTenantCdnType.Public, publicCdn.Enabled);
                        tenantContext.ExecuteQueryRetry();
                        if (publicCdn.Enabled)
                        {
                            ProcessOrigins(tenant, publicCdn, SPOTenantCdnType.Public, parser);
                            ProcessPolicies(tenant, publicCdn, SPOTenantCdnType.Public, parser);
                        }
                    }
                    var privateCdn = provisioningTenant.ContentDeliveryNetwork.PrivateCdn;
                    if(privateCdn != null)
                    {
                        tenant.SetTenantCdnEnabled(SPOTenantCdnType.Private, privateCdn.Enabled);
                        tenantContext.ExecuteQueryRetry();
                        if (privateCdn.Enabled)
                        {
                            ProcessOrigins(tenant, publicCdn, SPOTenantCdnType.Public, parser);
                            ProcessPolicies(tenant, publicCdn, SPOTenantCdnType.Public, parser);
                        }
                    }
                }
            }
        }

        private static void ProcessOrigins(Tenant tenant, CdnSettings cdnSettings, SPOTenantCdnType cdnType, TokenParser parser)
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
                                    tenant.AddPublicCdnOrigin(parsedOriginUrl);
                                }
                                break;
                            }
                        case OriginAction.Remove:
                            {
                                var parsedOriginUrl = parser.ParseString(origin.Url);
                                if (origins.Contains(parsedOriginUrl))
                                {
                                    tenant.RemovePublicCdnOrigin(parsedOriginUrl);
                                }
                                break;
                            }
                    }
                    tenant.Context.ExecuteQueryRetry();
                }
            }
        }

        private static void ProcessPolicies(Tenant tenant, CdnSettings cdnSettings, SPOTenantCdnType cdnType, TokenParser parser)
        {
            if(!string.IsNullOrEmpty(cdnSettings.IncludeFileExtensions))
            {
                var parsedValue = parser.ParseString(cdnSettings.IncludeFileExtensions);
                tenant.SetTenantCdnPolicy(cdnType, SPOTenantCdnPolicyType.IncludeFileExtensions, parsedValue);
            }
            if(!string.IsNullOrEmpty(cdnSettings.ExcludeRestrictedSiteClassifications))
            {
                var parsedValue = parser.ParseString(cdnSettings.ExcludeRestrictedSiteClassifications);
                tenant.SetTenantCdnPolicy(cdnType, SPOTenantCdnPolicyType.ExcludeRestrictedSiteClassifications, parsedValue);
            }
            if(!string.IsNullOrEmpty(cdnSettings.ExcludeIfNoScriptDisabled))
            {
                var parsedValue = parser.ParseString(cdnSettings.ExcludeIfNoScriptDisabled);
                tenant.SetTenantCdnPolicy(cdnType, SPOTenantCdnPolicyType.ExcludeIfNoScriptDisabled, parsedValue);
            }
            tenant.Context.ExecuteQueryRetry();
        }
    }

#endif
}
