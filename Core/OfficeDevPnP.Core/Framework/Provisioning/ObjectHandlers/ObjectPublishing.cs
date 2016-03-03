using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Xml.Linq;
using OfficeDevPnP.Core.Entities;
using System.IO;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.Web;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPublishing : ObjectContentHandlerBase
    {
        private const string AVAILABLEPAGELAYOUTS = "__PageLayouts";
        private const string DEFAULTPAGELAYOUT = "__DefaultPageLayout";
        private readonly Guid PUBLISHING_FEATURE_WEB = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
        private readonly Guid PUBLISHING_FEATURE_SITE = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");
        private const string PAGE_LAYOUT_CONTENT_TYPE_ID = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE811";
        private const string HTML_PAGE_LAYOUT_CONTENT_TYPE_ID = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE8110003D357F861E29844953D5CAA1D4D8A3B";
        private const string MASTER_PAGE_CONTENT_TYPE_ID = "0x010105";
        private const string HTML_MASTER_PAGE_CONTENT_TYPE_ID = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F0054AD6EF48B9F7B45A142F8173F171BD10003D357F861E29844953D5CAA1D4D8A3A";

        public override string Name
        {
            get { return "Publishing"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (web.IsFeatureActive(PUBLISHING_FEATURE_WEB))
                {
                    web.EnsureProperty(w => w.Language);
                    var webTemplates = web.GetAvailableWebTemplates(web.Language, false);
                    web.Context.Load(webTemplates, wts => wts.Include(wt => wt.Name, wt => wt.Lcid));
                    web.Context.ExecuteQueryRetry();
                    Publishing publishing = new Publishing();
                    publishing.AvailableWebTemplates.AddRange(webTemplates.AsEnumerable<WebTemplate>().Select(wt => new AvailableWebTemplate() { TemplateName = wt.Name, LanguageCode = (int)wt.Lcid }));
                    publishing.AutoCheckRequirements = AutoCheckRequirementsOptions.MakeCompliant;
                    publishing.DesignPackage = null;
                    publishing.PageLayouts.AddRange(GetAvailablePageLayouts(web));
                    template.Publishing = publishing;

                    ExtractMasterPagesAndPageLayouts(web, template, scope, creationInfo);
                }
            }
            return template;
        }

        private void ExtractMasterPagesAndPageLayouts(Web web, ProvisioningTemplate template, PnPMonitoredScope scope, ProvisioningTemplateCreationInformation creationInfo)
        {
            web.EnsureProperty(w => w.Url);
            String webApplicationUrl = GetWebApplicationUrl(web.Url);

            if (!String.IsNullOrEmpty(webApplicationUrl))
            {
                // Get the Publishing Feature reference template
                ProvisioningTemplate publishingFeatureTemplate = GetPublishingFeatureBaseTemplate();

                // Get a reference to the root folder of the master page gallery
                var gallery = web.GetCatalog(116);
                web.Context.Load(gallery, g => g.RootFolder);
                web.Context.ExecuteQueryRetry();

                var masterPageGalleryFolder = gallery.RootFolder;

                // Load the files in the master page gallery
                web.Context.Load(masterPageGalleryFolder.Files);
                web.Context.ExecuteQueryRetry();

                var sourceFiles = masterPageGalleryFolder.Files.AsEnumerable().Where(
                    f => f.Name.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase) ||
                    f.Name.EndsWith(".html", StringComparison.InvariantCultureIgnoreCase) ||
                    f.Name.EndsWith(".master", StringComparison.InvariantCultureIgnoreCase));

                foreach (var file in sourceFiles)
                {
                    var listItem = file.EnsureProperty(f => f.ListItemAllFields);
                    listItem.ContentType.EnsureProperties(ct => ct.Id, ct => ct.StringId);

                    // Check if the content type is of type Master Page or Page Layout
                    if (listItem.ContentType.StringId.StartsWith(MASTER_PAGE_CONTENT_TYPE_ID) ||
                        listItem.ContentType.StringId.StartsWith(PAGE_LAYOUT_CONTENT_TYPE_ID) ||
                        listItem.ContentType.StringId.StartsWith(HTML_MASTER_PAGE_CONTENT_TYPE_ID) ||
                        listItem.ContentType.StringId.StartsWith(HTML_PAGE_LAYOUT_CONTENT_TYPE_ID))
                    {
                        // Skip any .ASPX or .MASTER file related to an .HTML designer file
                        if ((file.Name.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase)
                            && sourceFiles.Any(f => f.Name.Equals(file.Name.ToLower().Replace(".aspx", ".html"), 
                                StringComparison.InvariantCultureIgnoreCase))) ||
                            (file.Name.EndsWith(".master", StringComparison.InvariantCultureIgnoreCase)
                            && sourceFiles.Any(f => f.Name.Equals(file.Name.ToLower().Replace(".master", ".html"),
                                StringComparison.InvariantCultureIgnoreCase))))
                        {
                            continue;
                        }

                        // If the file is a custom one, and not one native
                        // and coming out from the publishing feature
                        if (creationInfo.IncludeNativePublishingFiles || 
                            !IsPublishingFeatureNativeFile(publishingFeatureTemplate, file.Name))
                        {
                            var fullUri = new Uri(UrlUtility.Combine(webApplicationUrl, file.ServerRelativeUrl));

                            var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
                            var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

                            var publishingFile = new Model.File()
                            {
                                Folder = Tokenize(folderPath, web.Url),
                                Src = HttpUtility.UrlDecode(fileName),
                                Overwrite = true,
                            };

                            // Add field values to file
                            RetrieveFieldValues(web, file, publishingFile);

                            // Add the file to the template
                            template.Files.Add(publishingFile);

                            // Persist file using connector, if needed
                            if (creationInfo.PersistPublishingFiles)
                            {
                                PersistFile(web, creationInfo, scope, folderPath, fileName, true);
                            }

                            if (listItem.ContentType.StringId.StartsWith(MASTER_PAGE_CONTENT_TYPE_ID))
                            {
                                scope.LogWarning(String.Format("The file \"{0}\" is a custom MasterPage. Accordingly to the PnP Guidance (http://aka.ms/o365pnpguidancemasterpages) you should try to avoid using custom MasterPages.", file.Name));
                            }
                        }
                        else
                        {
                            scope.LogWarning(String.Format("Skipping file \"{0}\" because it is native in the publishing feature.", file.Name));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This method returns the reference template for publishing feature
        /// </summary>
        /// <returns>The reference template for publishing feature</returns>
        private ProvisioningTemplate GetPublishingFeatureBaseTemplate()
        {
            ProvisioningTemplate result = null;

            string nativeFilesTemplatePath = string.Format("OfficeDevPnP.Core.Framework.Provisioning.BaseTemplates.Common.Publishing-Feature-Native-Files.xml");
            using (Stream stream = typeof(BaseTemplateManager).Assembly.GetManifestResourceStream(nativeFilesTemplatePath))
            {
                // Figure out the formatter to use
                XDocument xTemplate = XDocument.Load(stream);
                var namespaceDeclarations = xTemplate.Root.Attributes().Where(a => a.IsNamespaceDeclaration).
                        GroupBy(a => a.Name.Namespace == XNamespace.None ? String.Empty : a.Name.LocalName,
                                a => XNamespace.Get(a.Value)).
                        ToDictionary(g => g.Key,
                                     g => g.First());
                var pnpns = namespaceDeclarations["pnp"];

                stream.Seek(0, SeekOrigin.Begin);

                // Get the XML document from the stream
                ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(pnpns.NamespaceName);

                // And convert it into a template
                result = formatter.ToProvisioningTemplate(stream);
            }

            return (result);
        }

        /// <summary>
        /// This method checks if the filename (for master pages and page layouts) 
        /// is native or custom for the publishing feature
        /// </summary>
        /// <param name="nativeFilesTemplate">The reference template for publishing feature</param>
        /// <param name="fileName">The filename to check</param>
        /// <returns>Whether the file is native or not for the publishing feature</returns>
        private Boolean IsPublishingFeatureNativeFile(ProvisioningTemplate nativeFilesTemplate, String fileName)
        {
            Boolean result = false;

            if (nativeFilesTemplate != null 
                && nativeFilesTemplate.Files != null 
                && nativeFilesTemplate.Files.Count > 0)
            {
                result = nativeFilesTemplate.Files.Any(f => f.Src == fileName);
            }

            return (result);
        }

        /// <summary>
        /// This method retrieves the Web Application URL of the provided site
        /// </summary>
        /// <param name="webUrl">The target web site URL</param>
        /// <returns>The Web Application URL</returns>
        private String GetWebApplicationUrl(String webUrl)
        {
            Uri uri = null;
            String result = null;
            if (Uri.TryCreate(webUrl, UriKind.Absolute, out uri))
            {
                result = string.Format("{0}://{1}/", uri.Scheme, uri.Authority);
            }

            return (result);
        }

        private IEnumerable<PageLayout> GetAvailablePageLayouts(Web web)
        {
            var defaultLayoutXml = web.GetPropertyBagValueString(DEFAULTPAGELAYOUT, null);

            var defaultPageLayoutUrl = string.Empty;
            if (defaultLayoutXml != null && defaultLayoutXml != "__inherit")
            {
                defaultPageLayoutUrl = XElement.Parse(defaultLayoutXml).Attribute("url").Value.Replace("_catalogs/masterpage/", String.Empty);
            }

            List<PageLayout> layouts = new List<PageLayout>();

            var layoutsXml = web.GetPropertyBagValueString(AVAILABLEPAGELAYOUTS, null);

            if (!string.IsNullOrEmpty(layoutsXml) && layoutsXml != "__inherit")
            {
                var layoutsElement = XElement.Parse(layoutsXml);

                foreach (var layout in layoutsElement.Descendants("layout"))
                {
                    if (layout.Attribute("url") != null)
                    {
                        var pageLayout = new PageLayout();
                        pageLayout.Path = layout.Attribute("url").Value.Replace("_catalogs/masterpage/", String.Empty);

                        if (pageLayout.Path == defaultPageLayoutUrl)
                        {
                            pageLayout.IsDefault = true;
                        }
                        layouts.Add(pageLayout);
                    }

                }
            }
            return layouts;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var site = (web.Context as ClientContext).Site;

                var webFeatureActive = web.IsFeatureActive(PUBLISHING_FEATURE_WEB);
                var siteFeatureActive = site.IsFeatureActive(PUBLISHING_FEATURE_SITE);
                if (template.Publishing.AutoCheckRequirements == AutoCheckRequirementsOptions.SkipIfNotCompliant && !webFeatureActive)
                {
                    scope.LogDebug("Publishing Feature (Web Scoped) not active. Skipping provisioning of Publishing settings");
                    return parser;
                }
                else if (template.Publishing.AutoCheckRequirements == AutoCheckRequirementsOptions.MakeCompliant)
                {
                    if (!siteFeatureActive)
                    {
                        scope.LogDebug("Making site compliant for publishing");
                        site.ActivateFeature(PUBLISHING_FEATURE_SITE);
                        web.ActivateFeature(PUBLISHING_FEATURE_WEB);
                    }
                    else
                    {
                        if (!web.IsFeatureActive(PUBLISHING_FEATURE_WEB))
                        {
                            scope.LogDebug("Making site compliant for publishing");
                            web.ActivateFeature(PUBLISHING_FEATURE_WEB);
                        }
                    }
                }
                else
                {
                    throw new Exception("Publishing Feature not active. Provisioning failed");
                }

                // Set allowed web templates
                var availableWebTemplates = template.Publishing.AvailableWebTemplates.Select(t => new WebTemplateEntity() { LanguageCode = t.LanguageCode.ToString(), TemplateName = t.TemplateName }).ToList();
                if (availableWebTemplates.Any())
                {
                    web.SetAvailableWebTemplates(availableWebTemplates);
                }

                // Set allowed page layouts
                var availablePageLayouts = template.Publishing.PageLayouts.Select(p => p.Path);
                if (availablePageLayouts.Any())
                {
                    web.SetAvailablePageLayouts(site.RootWeb, availablePageLayouts);
                }

                // Set default page layout, if any
                var defaultPageLayout = template.Publishing.PageLayouts.FirstOrDefault(p => p.IsDefault);
                if (defaultPageLayout != null)
                {
                    web.SetDefaultPageLayoutForSite(site.RootWeb, defaultPageLayout.Path);
                }

                if (template.Publishing.DesignPackage != null)
                {
                    var package = template.Publishing.DesignPackage;

                    var tempFileName = Path.Combine(Path.GetTempPath(), template.Connector.GetFilenamePart(package.DesignPackagePath));
                    scope.LogDebug("Saving {0} to temporary file: {1}", package.DesignPackagePath, tempFileName);
                    using (var stream = template.Connector.GetFileStream(package.DesignPackagePath))
                    {
                        using (var outstream = System.IO.File.Create(tempFileName))
                        {
                            stream.CopyTo(outstream);
                        }
                    }
                    scope.LogDebug("Installing design package");
                    site.InstallSolution(package.PackageGuid, tempFileName, package.MajorVersion, package.MinorVersion);
                    System.IO.File.Delete(tempFileName);
                }
                return parser;
            }
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return web.IsFeatureActive(PUBLISHING_FEATURE_WEB);
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return template.Publishing != null;
        }
    }
}
