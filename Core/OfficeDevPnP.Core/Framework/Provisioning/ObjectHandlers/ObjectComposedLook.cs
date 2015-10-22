using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.IO;
using OfficeDevPnP.Core.Diagnostics;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectComposedLook : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Composed Looks"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.ComposedLook != null &&
                    !template.ComposedLook.Equals(ComposedLook.Empty))
                {
                    bool executeQueryNeeded = false;

                    // Apply alternate CSS
                    if (!string.IsNullOrEmpty(template.ComposedLook.AlternateCSS))
                    {
                        var alternateCssUrl = parser.ParseString(template.ComposedLook.AlternateCSS);
                        web.AlternateCssUrl = alternateCssUrl;
                        web.Update();
                        executeQueryNeeded = true;
                    }

                    // Apply Site logo
                    if (!string.IsNullOrEmpty(template.ComposedLook.SiteLogo))
                    {
                        var siteLogoUrl = parser.ParseString(template.ComposedLook.SiteLogo);
                        web.SiteLogoUrl = siteLogoUrl;
                        web.Update();
                        executeQueryNeeded = true;
                    }

                    if (executeQueryNeeded)
                    {
                        web.Context.ExecuteQueryRetry();
                    }

                    if (String.IsNullOrEmpty(template.ComposedLook.ColorFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.FontFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                    {
                        // Apply OOB theme
                        web.SetComposedLookByUrl(template.ComposedLook.Name);
                    }
                    else
                    {
                        // Apply custom theme
                        string colorFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            colorFile = parser.ParseString(template.ComposedLook.ColorFile);
                        }
                        string backgroundFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            backgroundFile = parser.ParseString(template.ComposedLook.BackgroundFile);
                        }
                        string fontFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {
                            fontFile = parser.ParseString(template.ComposedLook.FontFile);
                        }

                        string masterUrl = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.MasterPage))
                        {
                            masterUrl = parser.ParseString(template.ComposedLook.MasterPage);
                        }
                        web.CreateComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl);
                        web.SetComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl);
                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Load object if not there
#if !CLIENTSDKV15
                web.EnsureProperties(w => w.Url, w => w.MasterUrl, w => w.AlternateCssUrl, w => w.SiteLogoUrl);
#else
                web.EnsureProperties(w => w.Url, w => w.MasterUrl);
#endif

                // Information coming from the site
                template.ComposedLook.MasterPage = Tokenize(web.MasterUrl, web.Url);
#if !CLIENTSDKV15
                template.ComposedLook.AlternateCSS = Tokenize(web.AlternateCssUrl, web.Url);
                template.ComposedLook.SiteLogo = Tokenize(web.SiteLogoUrl, web.Url);
#else
                template.ComposedLook.AlternateCSS = null;
                template.ComposedLook.SiteLogo = null;
#endif
                scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Retrieving_current_composed_look);
                var theme = web.GetCurrentComposedLook();


                if (theme != null)
                {
                    if (creationInfo != null)
                    {
                        // Don't exclude the DesignPreviewThemedCssFolderUrl property bag, if any
                        creationInfo.PropertyBagPropertiesToPreserve.Add("DesignPreviewThemedCssFolderUrl");
                    }

                    template.ComposedLook.Name = theme.Name;

                    if (theme.IsCustomComposedLook)
                    {
                        if (creationInfo != null && creationInfo.PersistComposedLookFiles && creationInfo.FileConnector != null)
                        {
                            Site site = (web.Context as ClientContext).Site;
                            if (!site.IsObjectPropertyInstantiated("Url"))
                            {
                                web.Context.Load(site);
                                web.Context.ExecuteQueryRetry();
                            }

                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Creating_SharePointConnector);
                            // Let's create a SharePoint connector since our files anyhow are in SharePoint at this moment
                            SharePointConnector spConnector = new SharePointConnector(web.Context, web.Url, "dummy");

                            // to get files from theme catalog we need a connector linked to the root site
                            SharePointConnector spConnectorRoot;
                            if (!site.Url.Equals(web.Url, StringComparison.InvariantCultureIgnoreCase))
                            {
                                spConnectorRoot = new SharePointConnector(web.Context.Clone(site.Url), site.Url, "dummy");
                            }
                            else
                            {
                                spConnectorRoot = spConnector;
                            }

                            // Download the theme/branding specific files
#if !CLIENTSDKV15
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, web.AlternateCssUrl, scope);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, web.SiteLogoUrl, scope);
#endif
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.BackgroundImage, scope);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.Theme, scope);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.Font, scope);
                        }

                        template.ComposedLook.BackgroundFile = FixFileUrl(Tokenize(theme.BackgroundImage, web.Url));
                        template.ComposedLook.ColorFile = FixFileUrl(Tokenize(theme.Theme, web.Url));
                        template.ComposedLook.FontFile = FixFileUrl(Tokenize(theme.Font, web.Url));

                        // Create file entries for the custom theme files  
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.BackgroundFile));
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.ColorFile));
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {

                            template.Files.Add(GetComposedLookFile(template.ComposedLook.FontFile));
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.SiteLogo))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.SiteLogo));
                        }

                        // If a base template is specified then use that one to "cleanup" the generated template model
                        if (creationInfo != null && creationInfo.BaseTemplate != null)
                        {
                            template = CleanupEntities(template, creationInfo.BaseTemplate);
                        }
                    }
                    else
                    {
                        template.ComposedLook.BackgroundFile = "";
                        template.ComposedLook.ColorFile = "";
                        template.ComposedLook.FontFile = "";
                    }
                }
                else
                {
                    template.ComposedLook = null;
                }

                if (creationInfo != null && creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private void DownLoadFile(SharePointConnector reader, SharePointConnector readerRoot, FileConnectorBase writer, string webUrl, string asset, PnPMonitoredScope scope)
        {

            // No file passed...leave
            if (String.IsNullOrEmpty(asset))
            {
                return;
            }

            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_DownLoadFile_Downloading_asset___0_, asset);
            ;

            SharePointConnector readerToUse;
            Model.File f = GetComposedLookFile(asset);

            // Strip the /sites/root part from /sites/root/lib/folder structure
            Uri u = new Uri(webUrl);
            if (f.Folder.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                f.Folder = f.Folder.Replace(u.PathAndQuery, "");
            }

            // in case of a theme catalog we need to use the root site reader as that list only exists on root site level
            if (f.Folder.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                readerToUse = readerRoot;
            }
            else
            {
                readerToUse = reader;
            }

            using (Stream s = readerToUse.GetFileStream(f.Src, f.Folder))
            {
                if (s != null)
                {
                    writer.SaveFileStream(f.Src, s);
                }
            }
        }

        private String FixFileName(string originalFileName)
        {
            // if we've found the file use the provided writer to persist the downloaded file
            String regexStrip = @"(\\|/|:|\*|\?|""|>|<|\||=)*";
            String result = Regex.Replace(originalFileName.Substring(0,
                originalFileName.IndexOf("?") > 0 ? originalFileName.IndexOf("?") : originalFileName.Length),
                regexStrip, "", RegexOptions.IgnorePatternWhitespace);

            return (result);
        }
        private String FixFileUrl(string originalFileUrl)
        {
            if (string.IsNullOrEmpty(originalFileUrl))
            {
                return "";
            }

            String fileUrl = originalFileUrl.Substring(0, originalFileUrl.LastIndexOf("/"));
            String fileName = FixFileName(originalFileUrl.Substring(originalFileUrl.LastIndexOf("/") + 1));

            String result = String.Format("{0}/{1}", fileUrl, fileName);

            return (result);
        }

        private Model.File GetComposedLookFile(string asset)
        {
            int index = asset.LastIndexOf("/");
            Model.File file = new Model.File();
            file.Src = FixFileName(asset.Substring(index + 1));
            file.Folder = asset.Substring(0, index);
            file.Overwrite = true;
            file.Security = null;
            return file;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            if (template.ComposedLook != null && baseTemplate.ComposedLook != null)
            {
                if (template.ComposedLook.Equals(baseTemplate.ComposedLook))
                {
                    template.ComposedLook = null;
                }

            }
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = (template.ComposedLook != null && !template.ComposedLook.Equals(ComposedLook.Empty));
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }
}
