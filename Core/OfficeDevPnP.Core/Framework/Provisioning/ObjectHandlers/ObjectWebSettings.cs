using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using OfficeDevPnP.Core.Utilities;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWebSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Web Settings"; }
        }

        public override string InternalName => "WebSettings";
        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(
#if !SP2013 && !SP2016
                    w => w.NoCrawl,
                    w => w.CommentsOnSitePagesDisabled,
                    w => w.ExcludeFromOfflineClient,
                    w => w.MembersCanShare,
                    w => w.DisableFlows,
                    w => w.DisableAppViews,
                    w => w.HorizontalQuickLaunch,
                    w => w.QuickLaunchEnabled,
#if !SP2019
                    w => w.SearchScope,
                    w => w.SearchBoxInNavBar,
    #endif
#endif
                    //w => w.Title,
                    //w => w.Description,
                    w => w.MasterUrl,
                    w => w.CustomMasterUrl,
                    w => w.SiteLogoUrl,
                    w => w.RequestAccessEmail,
                    w => w.RootFolder,
                    w => w.AlternateCssUrl,
                    w => w.ServerRelativeUrl,
                    w => w.Url
                    );

                var webSettings = new WebSettings();
#if !SP2013 && !SP2016
                webSettings.NoCrawl = web.NoCrawl;
                webSettings.CommentsOnSitePagesDisabled = web.CommentsOnSitePagesDisabled;
                webSettings.ExcludeFromOfflineClient = web.ExcludeFromOfflineClient;
                webSettings.MembersCanShare = web.MembersCanShare;
                webSettings.DisableFlows = web.DisableFlows;
                webSettings.DisableAppViews = web.DisableAppViews;
                webSettings.HorizontalQuickLaunch = web.HorizontalQuickLaunch;
                webSettings.QuickLaunchEnabled = web.QuickLaunchEnabled;
#if !SP2019
                webSettings.SearchScope = (SearchScopes)Enum.Parse(typeof(SearchScopes), web.SearchScope.ToString(), true);
                webSettings.SearchBoxInNavBar = (SearchBoxInNavBar)Enum.Parse(typeof(SearchBoxInNavBar), web.SearchBoxInNavBar.ToString(), true);
                webSettings.SearchCenterUrl = web.GetWebSearchCenterUrl(true);
    #endif
#endif
                // We're not extracting Title and Description
                //webSettings.Title = Tokenize(web.Title, web.Url);
                //webSettings.Description = Tokenize(web.Description, web.Url);
                webSettings.MasterPageUrl = Tokenize(web.MasterUrl, web.Url);
                webSettings.CustomMasterPageUrl = Tokenize(web.CustomMasterUrl, web.Url);
                webSettings.SiteLogo = TokenizeHost(web, Tokenize(web.SiteLogoUrl, web.Url));
                // Notice. No tokenization needed for the welcome page, it's always relative for the site
                webSettings.WelcomePage = web.RootFolder.WelcomePage;
                webSettings.AlternateCSS = Tokenize(web.AlternateCssUrl, web.Url);
                webSettings.RequestAccessEmail = web.RequestAccessEmail;

#if !ONPREMISES
                // Can we get the hubsite url? This requires Tenant Admin rights
                try
                {
                    var site = ((ClientContext)web.Context).Site;
                    site.EnsureProperties(s => s.HubSiteId, s => s.Id);
                    if (site.HubSiteId != Guid.Empty && site.HubSiteId != site.Id)
                    {
                        using (var tenantContext = web.Context.Clone((web.Context as ClientContext).Web.GetTenantAdministrationUrl()))
                        {
                            var tenant = new Tenant(tenantContext);
                            var hubsiteProperties = tenant.GetHubSitePropertiesById(site.HubSiteId);
                            tenantContext.Load(hubsiteProperties);
                            tenantContext.ExecuteQueryRetry();
                            webSettings.HubSiteUrl = hubsiteProperties.SiteUrl;
                        }
                    }
                }
                catch { }
#endif

                if (creationInfo.PersistBrandingFiles)
                {
                    if (!string.IsNullOrEmpty(web.MasterUrl))
                    {
                        var masterUrl = web.MasterUrl.ToLower();
                        if (!masterUrl.EndsWith("default.master") && !masterUrl.EndsWith("custom.master") && !masterUrl.EndsWith("v4.master") && !masterUrl.EndsWith("seattle.master") && !masterUrl.EndsWith("oslo.master"))
                        {

                            if (PersistFile(web, creationInfo, scope, web.MasterUrl))
                            {
                                template.Files.Add(GetTemplateFile(web, web.MasterUrl));
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(web.CustomMasterUrl))
                    {
                        var customMasterUrl = web.CustomMasterUrl.ToLower();
                        if (!customMasterUrl.EndsWith("default.master") && !customMasterUrl.EndsWith("custom.master") && !customMasterUrl.EndsWith("v4.master") && !customMasterUrl.EndsWith("seattle.master") && !customMasterUrl.EndsWith("oslo.master"))
                        {
                            if (PersistFile(web, creationInfo, scope, web.CustomMasterUrl))
                            {
                                template.Files.Add(GetTemplateFile(web, web.CustomMasterUrl));
                            }
                        }
                    }
                    // Extract site logo if property has been set and it's not dynamic image from _api URL
                    if (!string.IsNullOrEmpty(web.SiteLogoUrl) && (!web.SiteLogoUrl.ToLowerInvariant().Contains("_api/")))
                    {
                        // Convert to server relative URL if needed (web.SiteLogoUrl can be set to FQDN URL of a file hosted in the site (e.g. for communication sites))
                        Uri webUri = new Uri(web.Url);
                        string webUrl = $"{webUri.Scheme}://{webUri.DnsSafeHost}";
                        string siteLogoServerRelativeUrl = web.SiteLogoUrl.Replace(webUrl, "");

                        if (PersistFile(web, creationInfo, scope, siteLogoServerRelativeUrl))
                        {
                            template.Files.Add(GetTemplateFile(web, HttpUtility.UrlDecode(siteLogoServerRelativeUrl)));
                        }
                    }
                    if (!string.IsNullOrEmpty(web.AlternateCssUrl))
                    {
                        if (PersistFile(web, creationInfo, scope, web.AlternateCssUrl))
                        {
                            template.Files.Add(GetTemplateFile(web, web.AlternateCssUrl));
                        }
                    }
                    var files = template.Files.Distinct().ToList();
                    template.Files.Clear();
                    template.Files.AddRange(files);
                }

                if (!creationInfo.PersistBrandingFiles)
                {
                    if (creationInfo.BaseTemplate != null)
                    {
                        if (!webSettings.Equals(creationInfo.BaseTemplate.WebSettings))
                        {
                            template.WebSettings = webSettings;
                        }
                    }
                    else
                    {
                        template.WebSettings = webSettings;
                    }
                }
                else
                {
                    template.WebSettings = webSettings;
                }
            }
            return template;
        }

        //TODO: Candidate for cleanup
        private Model.File GetTemplateFile(Web web, string serverRelativeUrl)
        {

            var webServerUrl = web.EnsureProperty(w => w.Url);
            var serverUri = new Uri(webServerUrl);
            var serverUrl = $"{serverUri.Scheme}://{serverUri.Authority}";
            var fullUri = new Uri(UrlUtility.Combine(serverUrl, serverRelativeUrl));

            var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
            var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

            // store as site relative path
            folderPath = folderPath.Replace(web.ServerRelativeUrl, "").Trim('/');
            var templateFile = new Model.File()
            {
                Folder = Tokenize(folderPath, web.Url),
                Src = !string.IsNullOrEmpty(folderPath) ? $"{folderPath}/{fileName}" : fileName,
                Overwrite = true,
            };

            return templateFile;
        }

        private bool PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string serverRelativeUrl)
        {
            var success = false;
            if (creationInfo.PersistBrandingFiles)
            {
                if (creationInfo.FileConnector != null)
                {
                    if (UrlUtility.IsIisVirtualDirectory(serverRelativeUrl))
                    {
                        scope.LogWarning("File is not located in the content database. Not retrieving {0}", serverRelativeUrl);
                        return success;
                    }

                    try
                    {
                        var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
                        string fileName = string.Empty;
                        if (serverRelativeUrl.IndexOf("/") > -1)
                        {
                            fileName = serverRelativeUrl.Substring(serverRelativeUrl.LastIndexOf("/") + 1);
                        }
                        else
                        {
                            fileName = serverRelativeUrl;
                        }
                        web.Context.Load(file);
                        web.Context.ExecuteQueryRetry();
                        ClientResult<Stream> stream = file.OpenBinaryStream();
                        web.Context.ExecuteQueryRetry();

                        var baseUri = new Uri(web.Url);
                        var fullUri = new Uri(baseUri, file.ServerRelativeUrl);
                        var folderPath = HttpUtility.UrlDecode(fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));

                        // Configure the filename to use 
                        fileName = HttpUtility.UrlDecode(fullUri.Segments[fullUri.Segments.Count() - 1]);

                        // Build up a site relative container URL...might end up empty as well
                        String container = HttpUtility.UrlDecode(folderPath.Replace(web.ServerRelativeUrl, "")).Trim('/').Replace("/", "\\");

                        using (Stream memStream = new MemoryStream())
                        {
                            CopyStream(stream.Value, memStream);
                            memStream.Position = 0;
                            if (!string.IsNullOrEmpty(container))
                            {
                                creationInfo.FileConnector.SaveFileStream(fileName, container, memStream);
                            }
                            else
                            {
                                creationInfo.FileConnector.SaveFileStream(fileName, memStream);
                            }
                        }
                        success = true;
                    }
                    catch (ServerException ex1)
                    {
                        // If we are referring a file from a location outside of the current web or at a location where we cannot retrieve the file an exception is thrown. We swallow this exception.
                        if (ex1.ServerErrorCode != -2147024809)
                        {
                            throw;
                        }
                        else
                        {
                            scope.LogWarning("File is not necessarily located in the current web. Not retrieving {0}", serverRelativeUrl);
                        }
                    }
                }
                else
                {
                    WriteMessage("No connector present to persist site logo.", ProvisioningMessageType.Error);
                    scope.LogError("No connector present to persist site logo");
                }
            }
            else
            {
                success = true;
            }
            return success;
        }

        private string TokenizeHost(Web web, string json)
        {
            // HostUrl token replacement
            var uri = new Uri(web.Url);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}:{uri.Port}", "{hosturl}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}", "{hosturl}", RegexOptions.IgnoreCase);
            return json;
        }


        private void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.WebSettings != null)
                {
                    // Check if this is not a noscript site as we're not allowed to update some properties
                    bool isNoScriptSite = web.IsNoScriptSite();

                    web.EnsureProperties(
#if !SP2013 && !SP2016
                        w => w.NoCrawl,
                        w => w.CommentsOnSitePagesDisabled,
                        w => w.ExcludeFromOfflineClient,
                        w => w.MembersCanShare,
                        w => w.DisableFlows,
                        w => w.DisableAppViews,
                        w => w.HorizontalQuickLaunch,
#if !SP2019
                        w => w.SearchScope,
                        w => w.SearchBoxInNavBar,
#endif
#endif
                        w => w.WebTemplate,
                        w => w.HasUniqueRoleAssignments);

                    var webSettings = template.WebSettings;

                    // Since the IsSubSite function can trigger an executequery ensure it's called before any updates to the web object are done.
                    if (!web.IsSubSite() || (web.IsSubSite() && web.HasUniqueRoleAssignments))
                    {
                        String requestAccessEmailValue = parser.ParseString(webSettings.RequestAccessEmail);
                        if (!String.IsNullOrEmpty(requestAccessEmailValue) && requestAccessEmailValue.Length >= 255)
                        {
                            requestAccessEmailValue = requestAccessEmailValue.Substring(0, 255);
                        }
                        if (!String.IsNullOrEmpty(requestAccessEmailValue))
                        {
                            web.RequestAccessEmail = requestAccessEmailValue;

                            web.Update();
                            web.Context.ExecuteQueryRetry();
                        }
                    }

#if !SP2013 && !SP2016
                    if (!isNoScriptSite)
                    {
                        web.NoCrawl = webSettings.NoCrawl;
                    }
                    else
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_WebSettings_SkipNoCrawlUpdate);
                    }

                    if (web.CommentsOnSitePagesDisabled != webSettings.CommentsOnSitePagesDisabled)
                    {
                        web.CommentsOnSitePagesDisabled = webSettings.CommentsOnSitePagesDisabled;
                    }

                    if (web.ExcludeFromOfflineClient != webSettings.ExcludeFromOfflineClient)
                    {
                        web.ExcludeFromOfflineClient = webSettings.ExcludeFromOfflineClient;
                    }

                    if (web.MembersCanShare != webSettings.MembersCanShare)
                    {
                        web.MembersCanShare = webSettings.MembersCanShare;
                    }

                    if (web.DisableFlows != webSettings.DisableFlows)
                    {
                        web.DisableFlows = webSettings.DisableFlows;
                    }

                    if (web.DisableAppViews != webSettings.DisableAppViews)
                    {
                        web.DisableAppViews = webSettings.DisableAppViews;
                    }

                    if (web.HorizontalQuickLaunch != webSettings.HorizontalQuickLaunch)
                    {
                        web.HorizontalQuickLaunch = webSettings.HorizontalQuickLaunch;
                    }

#if !SP2019
                    if (web.SearchScope.ToString() != webSettings.SearchScope.ToString())
                    {
                        web.SearchScope = (SearchScopeType)Enum.Parse(typeof(SearchScopeType), webSettings.SearchScope.ToString(), true);
                    }

                    if(web.SearchBoxInNavBar.ToString() != webSettings.SearchBoxInNavBar.ToString())
                    {
                        web.SearchBoxInNavBar = (SearchBoxInNavBarType)Enum.Parse(typeof(SearchBoxInNavBarType), webSettings.SearchBoxInNavBar.ToString(), true);
                    }

                    if (!string.IsNullOrEmpty(webSettings.SearchCenterUrl) &&
                        web.GetWebSearchCenterUrl(true) != webSettings.SearchCenterUrl)
                    {
                        web.SetWebSearchCenterUrl(webSettings.SearchCenterUrl);
                    }
#endif
#endif
                    var masterUrl = parser.ParseString(webSettings.MasterPageUrl);
                    if (!string.IsNullOrEmpty(masterUrl))
                    {
                        if (!isNoScriptSite)
                        {
                            web.MasterUrl = masterUrl;
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_WebSettings_SkipMasterPageUpdate);
                        }
                    }
                    var customMasterUrl = parser.ParseString(webSettings.CustomMasterPageUrl);
                    if (!string.IsNullOrEmpty(customMasterUrl))
                    {
                        if (!isNoScriptSite)
                        {
                            web.CustomMasterUrl = customMasterUrl;
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_WebSettings_SkipCustomMasterPageUpdate);
                        }
                    }
                    if (webSettings.Title != null)
                    {
                        web.Title = parser.ParseString(webSettings.Title);
                    }
                    if (webSettings.Description != null)
                    {
                        web.Description = parser.ParseString(webSettings.Description);
                    }
                    if (webSettings.SiteLogo != null)
                    {
                        var logoUrl = parser.ParseString(webSettings.SiteLogo);
                        if (template.BaseSiteTemplate == "SITEPAGEPUBLISHING#0" && web.WebTemplate == "GROUP")
                        {
                            // logo provisioning throws when applying across base template IDs; provisioning fails in this case
                            // this is the error that is already (rightly so) shown beforehand in the console: WARNING: The source site from which the template was generated had a base template ID value of SITEPAGEPUBLISHING#0, while the current target site has a base template ID value of GROUP#0. This could cause potential issues while applying the template.
                            WriteMessage("Applying site logo across base template IDs is not possible. Skipping site logo provisioning.", ProvisioningMessageType.Warning);
                        }
                        else
                        // Modern site? Then we assume the SiteLogo is actually a filepath
                        if (web.WebTemplate == "GROUP")
                        {
#if !ONPREMISES
                            if (!string.IsNullOrEmpty(logoUrl) && !logoUrl.ToLower().Contains("_api/groupservice/getgroupimage"))
                            {
                                var fileBytes = ConnectorFileHelper.GetFileBytes(template.Connector, logoUrl);
                                if (fileBytes != null && fileBytes.Length > 0)
                                {
#if !NETSTANDARD2_0
                                    var mimeType = MimeMapping.GetMimeMapping(logoUrl);
#else
                                    var mimeType = "";
                                    var imgUrl = logoUrl;
                                    if (imgUrl.Contains("?"))
                                    {
                                        imgUrl = imgUrl.Split(new[] { '?' })[0];
                                    }
                                    if(imgUrl.EndsWith(".gif",StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        mimeType = "image/gif";
                                    }
                                    if (imgUrl.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        mimeType = "image/png";
                                    }
                                    if (imgUrl.EndsWith(".jpg", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        mimeType = "image/jpeg";
                                    }
#endif
                                    Sites.SiteCollection.SetGroupImageAsync((ClientContext)web.Context, fileBytes, mimeType).GetAwaiter().GetResult();

                                }
                            }
#endif
                        }
                        else
                        {
                            web.SiteLogoUrl = logoUrl;
                        }
                    }
                    var welcomePage = parser.ParseString(webSettings.WelcomePage);
                    if (!string.IsNullOrEmpty(welcomePage))
                    {
                        web.RootFolder.WelcomePage = welcomePage;
                        web.RootFolder.Update();
                    }
                    if (webSettings.AlternateCSS != null)
                    {
                        web.AlternateCssUrl = parser.ParseString(webSettings.AlternateCSS);
                    }

                    // Tempory disabled as this change is a breaking change for folks that have not set this property in their provisioning templates
                    //web.QuickLaunchEnabled = webSettings.QuickLaunchEnabled;

                    web.Update();
                    web.Context.ExecuteQueryRetry();

#if !ONPREMISES
                    if (webSettings.HubSiteUrl != null)
                    {
                        var hubsiteUrl = parser.ParseString(webSettings.HubSiteUrl);
                        try
                        {
                            using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl(), applyingInformation.AccessTokens))
                            {
                                var tenant = new Tenant(tenantContext);
                                tenant.ConnectSiteToHubSite(web.Url, hubsiteUrl);
                                tenantContext.ExecuteQueryRetry();
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteMessage($"Hub site association failed: {ex.Message}", ProvisioningMessageType.Warning);
                        }
                    }
#endif
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return template.WebSettings != null;
        }


    }
}
