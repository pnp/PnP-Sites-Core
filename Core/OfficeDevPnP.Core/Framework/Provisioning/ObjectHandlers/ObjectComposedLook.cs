using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.IO;
using OfficeDevPnP.Core.Diagnostics;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectComposedLook : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Composed Looks"; }
        }

        public override string InternalName => "ComposedLooks";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.ComposedLook != null &&
                    !template.ComposedLook.IsEmptyOrBlank())
                {
                    // Check if this is not a noscript site as themes and composed looks are not supported
                    if (web.IsNoScriptSite())
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_NoSiteCheck);
                        return parser;
                    }

                    bool executeQueryNeeded = false;
                    if (executeQueryNeeded)
                    {
                        web.Context.ExecuteQueryRetry();
                    }

                    if (String.IsNullOrEmpty(template.ComposedLook.ColorFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.FontFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                    {
                        // Apply OOB theme
                        web.SetComposedLookByUrl(template.ComposedLook.Name, "", "", "");
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
                        if (template.WebSettings != null && !string.IsNullOrEmpty(template.WebSettings.MasterPageUrl))
                        {
                            masterUrl = parser.ParseString(template.WebSettings.MasterPageUrl);
                        }
                        web.CreateComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl);
                        web.SetComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl);
                    }

                    // Persist composed look info in property bag
                    var composedLookJson = JsonConvert.SerializeObject(template.ComposedLook);
                    web.SetPropertyBagValue("_PnP_ProvisioningTemplateComposedLookInfo", composedLookJson);
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Retrieving_current_composed_look);

                // Ensure that we have URL property loaded for web and site
                web.EnsureProperty(w => w.Url);
                Site site = (web.Context as ClientContext).Site;
                site.EnsureProperty(s => s.Url);
               
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

                // Check if we have composed look info in the property bag, if so, use that, otherwise try to detect the current composed look
                if (web.PropertyBagContainsKey("_PnP_ProvisioningTemplateComposedLookInfo"))
                {
                    scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Using_ComposedLookInfoFromPropertyBag);

                    try
                    {
                        var composedLook = JsonConvert.DeserializeObject<ComposedLook>(web.GetPropertyBagValueString("_PnP_ProvisioningTemplateComposedLookInfo", ""));
                        if (composedLook.Name == null)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_ComposedLookInfoFailedToDeserialize);
                            throw new JsonSerializationException();
                        }

                        composedLook.BackgroundFile = Tokenize(composedLook.BackgroundFile, web.Url);
                        composedLook.FontFile = Tokenize(composedLook.FontFile, web.Url);
                        composedLook.ColorFile = Tokenize(composedLook.ColorFile, web.Url);
                        template.ComposedLook = composedLook;

                        if (!web.IsSubSite() && creationInfo != null && 
                                creationInfo.PersistBrandingFiles && creationInfo.FileConnector != null)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Creating_SharePointConnector);
                            // Let's create a SharePoint connector since our files anyhow are in SharePoint at this moment
                            TokenParser parser = new TokenParser(web, template);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, parser.ParseString(composedLook.BackgroundFile), scope);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, parser.ParseString(composedLook.ColorFile), scope);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, parser.ParseString(composedLook.FontFile), scope);
                        }
                        // Create file entries for the custom theme files  
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            var f = GetComposedLookFile(template.ComposedLook.BackgroundFile);
                            f.Folder = Tokenize(f.Folder, web.Url);
                            template.Files.Add(f);
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            var f = GetComposedLookFile(template.ComposedLook.ColorFile);
                            f.Folder = Tokenize(f.Folder, web.Url);
                            template.Files.Add(f);
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {
                            var f = GetComposedLookFile(template.ComposedLook.FontFile);
                            f.Folder = Tokenize(f.Folder, web.Url);
                            template.Files.Add(f);
                        }

                    }
                    catch (JsonSerializationException)
                    {
                        // cannot deserialize the object, fall back to composed look detection
                        template = DetectComposedLook(web, template, creationInfo, scope, spConnector, spConnectorRoot);
                    }

                }
                else
                {
                    template = DetectComposedLook(web, template, creationInfo, scope, spConnector, spConnectorRoot);
                }

                if (creationInfo != null && creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate DetectComposedLook(Web web, ProvisioningTemplate template, 
                                                    ProvisioningTemplateCreationInformation creationInfo, 
                                                    PnPMonitoredScope scope, SharePointConnector spConnector, 
                                                    SharePointConnector spConnectorRoot)
        {

            var theme = web.GetCurrentComposedLook();

            if (theme != null)
            {
                if (creationInfo != null)
                {
                    // Don't exclude the DesignPreviewThemedCssFolderUrl property bag, if any
                    creationInfo.PropertyBagPropertiesToPreserve.Add("DesignPreviewThemedCssFolderUrl");
                }

                template.ComposedLook.Name = 
                    theme.Name != null ? theme.Name : String.Empty;

                if (theme.IsCustomComposedLook)
                {
                    // Set the URL pointers to files
                    template.ComposedLook.BackgroundFile = FixFileUrl(Tokenize(theme.BackgroundImage, web.Url));
                    template.ComposedLook.ColorFile = FixFileUrl(Tokenize(theme.Theme, web.Url));
                    template.ComposedLook.FontFile = FixFileUrl(Tokenize(theme.Font, web.Url));

                    // Download files if this is root site, since theme files are only stored there
                    if (!web.IsSubSite() && creationInfo != null && 
                        creationInfo.PersistBrandingFiles && creationInfo.FileConnector != null)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Creating_SharePointConnector);
                        // Let's create a SharePoint connector since our files anyhow are in SharePoint at this moment
                        // Download the theme/branding specific files
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.BackgroundImage, scope);
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.Theme, scope);
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.Font, scope);
                    }

                    // Create file entries for the custom theme files, but only if it's a root site
                    // If it's root site we do not extract or set theme files, since those are in the root of the site collection
                    if (!web.IsSubSite())
                    {   
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

            // Strip the /sites/root part from /sites/root/lib/folder structure, special case for root site handling.
            Uri u = new Uri(webUrl);
            if (f.Folder.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1 && u.PathAndQuery.Length > 1)
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

            String result = $"{fileUrl}/{fileName}";

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

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = (template.ComposedLook != null && !template.ComposedLook.IsEmptyOrBlank() && !web.IsNoScriptSite());
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = !web.IsNoScriptSite();
            }
            return _willExtract.Value;
        }
    }
}
