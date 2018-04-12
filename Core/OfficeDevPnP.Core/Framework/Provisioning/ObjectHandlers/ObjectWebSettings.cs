using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.IO;
using OfficeDevPnP.Core.Utilities;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWebSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Web Settings"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(
#if !ONPREMISES
                    w => w.NoCrawl,
                    w => w.RequestAccessEmail,
                    w => w.CommentsOnSitePagesDisabled,
#endif
                    //w => w.Title,
                    //w => w.Description,
                    w => w.MasterUrl,
                    w => w.CustomMasterUrl,
                    w => w.SiteLogoUrl,
                    w => w.RootFolder,
                    w => w.AlternateCssUrl,
                    w => w.ServerRelativeUrl,
                    w => w.Url);

                var webSettings = new WebSettings();
#if !ONPREMISES
                webSettings.NoCrawl = web.NoCrawl;
                webSettings.RequestAccessEmail = web.RequestAccessEmail;
                webSettings.CommentsOnSitePagesDisabled = web.CommentsOnSitePagesDisabled;
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
                            template.Files.Add(GetTemplateFile(web, siteLogoServerRelativeUrl));
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
                    WriteMessage("No connector present to persist homepage.", ProvisioningMessageType.Error);
                    scope.LogError("No connector present to persist homepage");
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
#if !ONPREMISES
                        w => w.CommentsOnSitePagesDisabled,
#endif
                        w => w.HasUniqueRoleAssignments);

                    var webSettings = template.WebSettings;
#if !ONPREMISES
                    if (!isNoScriptSite)
                    {
                        web.NoCrawl = webSettings.NoCrawl;
                    }
                    else
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_WebSettings_SkipNoCrawlUpdate);
                    }

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

                    if(web.CommentsOnSitePagesDisabled != webSettings.CommentsOnSitePagesDisabled)
                    {
                        web.CommentsOnSitePagesDisabled = webSettings.CommentsOnSitePagesDisabled;
                    }
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
                        web.SiteLogoUrl = parser.ParseString(webSettings.SiteLogo);
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
                    web.Update();
                    web.Context.ExecuteQueryRetry();

#if !ONPREMISES
                    if (webSettings.HubSiteUrl != null)
                    {
                        var hubsiteUrl = parser.ParseString(webSettings.HubSiteUrl);
                        try
                        {
                            using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl()))
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
