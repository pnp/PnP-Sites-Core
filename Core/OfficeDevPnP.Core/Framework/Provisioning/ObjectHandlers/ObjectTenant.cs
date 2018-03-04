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
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Tenant != null)
                {
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
    }
#endif
}
