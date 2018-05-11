using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPages : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Pages"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.RootFolder.WelcomePage);

                // Check if this is not a noscript site as we're not allowed to update some properties
                bool isNoScriptSite = web.IsNoScriptSite();

                foreach (var page in template.Pages)
                {
                    var url = parser.ParseString(page.Url);

                    if (!url.ToLower().StartsWith(web.ServerRelativeUrl.ToLower()))
                    {
                        url = UrlUtility.Combine(web.ServerRelativeUrl, url);
                    }

                    var exists = true;
                    Microsoft.SharePoint.Client.File file = null;
                    try
                    {
                        file = web.GetFileByServerRelativeUrl(url);
                        web.Context.Load(file);
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (ServerException ex)
                    {
                        if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                        {
                            exists = false;
                        }
                    }
                    if (exists)
                    {
                        if (page.Overwrite)
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0_, url);

                                // determine url of current home page
                                string welcomePageUrl = web.RootFolder.WelcomePage;
                                string welcomePageServerRelativeUrl = welcomePageUrl != null
                                    ? UrlUtility.Combine(web.ServerRelativeUrl, web.RootFolder.WelcomePage)
                                    : null;

                                bool overwriteWelcomePage = string.Equals(url, welcomePageServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase);

                                // temporarily reset home page so we can delete it
                                if (overwriteWelcomePage)
                                {
                                    web.SetHomePage(string.Empty);
                                }

                                file.DeleteObject();
                                web.Context.ExecuteQueryRetry();
                                web.AddWikiPageByUrl(url);
                                if (page.Layout == WikiPageLayout.Custom)
                                {
                                    web.AddLayoutToWikiPage(WikiPageLayout.OneColumn, url);
                                }
                                else
                                {
                                    web.AddLayoutToWikiPage(page.Layout, url);
                                }

                                if (overwriteWelcomePage)
                                {
                                    // restore welcome page to previous value
                                    web.SetHomePage(welcomePageUrl);
                                }
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0__failed___1_____2_, url, ex.Message, ex.StackTrace);
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Pages_Creating_new_page__0_, url);

                            web.AddWikiPageByUrl(url);
                            if (page.Layout == WikiPageLayout.Custom)
                            {
                                web.AddLayoutToWikiPage(WikiPageLayout.OneColumn, url);
                            }
                            else
                            {
                                web.AddLayoutToWikiPage(page.Layout, url);
                            }
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Creating_new_page__0__failed___1_____2_, url, ex.Message, ex.StackTrace);
                        }
                    }

#pragma warning disable 618
                    if (page.WelcomePage)
#pragma warning restore 618
                    {
                        web.RootFolder.EnsureProperty(p => p.ServerRelativeUrl);
                        var rootFolderRelativeUrl = url.Substring(web.RootFolder.ServerRelativeUrl.Length);
                        web.SetHomePage(rootFolderRelativeUrl);
                    }

#if !SP2013
                    bool webPartsNeedLocalization = false;
#endif
                    if (page.WebParts != null & page.WebParts.Any())
                    {
                        if (!isNoScriptSite)
                        {
                            var existingWebParts = web.GetWebParts(url);

                            foreach (var webPart in page.WebParts)
                            {
                                if (existingWebParts.FirstOrDefault(w => w.WebPart.Title == parser.ParseString(webPart.Title)) == null)
                                {
                                    WebPartEntity wpEntity = new WebPartEntity();
                                    wpEntity.WebPartTitle = parser.ParseString(webPart.Title);
                                    wpEntity.WebPartXml = parser.ParseXmlStringWebpart(webPart.Contents.Trim(new[] { '\n', ' ' }), web, "~sitecollection", "~site");
                                    var wpd = web.AddWebPartToWikiPage(url, wpEntity, (int)webPart.Row, (int)webPart.Column, false);
#if !SP2013
                                    if (webPart.Title.ContainsResourceToken())
                                    {
                                        // update data based on where it was added - needed in order to localize wp title
#if !SP2016
                                        wpd.EnsureProperties(w => w.ZoneId, w => w.WebPart, w => w.WebPart.Properties);
                                        webPart.Zone = wpd.ZoneId;
#else
                                        wpd.EnsureProperties(w => w.WebPart, w => w.WebPart.Properties);
#endif
                                        webPart.Order = (uint)wpd.WebPart.ZoneIndex;
                                        webPartsNeedLocalization = true;
                                    }
#endif
                                }
                            }

                            // Remove any existing WebPartIdToken tokens in the parser that were added by other pages. They won't apply to this page,
                            // and they'll cause issues if this page contains web parts with the same name as web parts on other pages.
                            parser.Tokens.RemoveAll(t => t is WebPartIdToken);

                            var allWebParts = web.GetWebParts(url);
                            foreach (var webpart in allWebParts)
                            {
                                parser.AddToken(new WebPartIdToken(web, webpart.WebPart.Title, webpart.Id));
                            }
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_Pages_SkipAddingWebParts, page.Url);
                        }
                    }

#if !SP2013
                    if (webPartsNeedLocalization)
                    {
                        page.LocalizeWebParts(web, parser, scope);
                    }
#endif

                    file = web.GetFileByServerRelativeUrl(url);
                    file.EnsureProperty(f => f.ListItemAllFields);

                    if (page.Fields.Any())
                    {
                        var item = file.ListItemAllFields;
                        foreach (var fieldValue in page.Fields)
                        {
                            item[fieldValue.Key] = parser.ParseString(fieldValue.Value);
                        }
                        item.Update();
                        web.Context.ExecuteQueryRetry();
                    }
                    if (page.Security != null && page.Security.RoleAssignments.Count != 0)
                    {
                        web.Context.Load(file.ListItemAllFields);
                        web.Context.ExecuteQueryRetry();
                        file.ListItemAllFields.SetSecurity(parser, page.Security);
                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Impossible to return all files in the site currently

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }


        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Pages.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }
}
