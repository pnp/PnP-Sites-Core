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
                        web.Context.ExecuteQuery();
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

                                if (page.WelcomePage && url.Contains(web.RootFolder.WelcomePage))
                                    web.SetHomePage(string.Empty);

                                file.DeleteObject();
                                web.Context.ExecuteQueryRetry();
                                web.AddWikiPageByUrl(url);
                                if (page.Layout == WikiPageLayout.Custom)
                                {
                                    web.AddLayoutToWikiPage(WikiPageLayout.OneColumn, url);
                                }
                                else {
                                    web.AddLayoutToWikiPage(page.Layout, url);
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
                            web.AddLayoutToWikiPage(page.Layout, url);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Creating_new_page__0__failed___1_____2_, url, ex.Message, ex.StackTrace);
                        }
                    }

                    if (page.WelcomePage)
                    {
                        web.RootFolder.EnsureProperty(p => p.ServerRelativeUrl);
                        var rootFolderRelativeUrl = url.Substring(web.RootFolder.ServerRelativeUrl.Length);
                        web.SetHomePage(rootFolderRelativeUrl);
                    }

                    if (page.WebParts != null & page.WebParts.Any())
                    {
                        var existingWebParts = web.GetWebParts(url);

                        foreach (var webpart in page.WebParts)
                        {
                            if (existingWebParts.FirstOrDefault(w => w.WebPart.Title == webpart.Title) == null)
                            {
                                WebPartEntity wpEntity = new WebPartEntity();
                                wpEntity.WebPartTitle = webpart.Title;
                                wpEntity.WebPartXml = parser.ParseString(webpart.Contents.Trim(new[] { '\n', ' ' }));
                                web.AddWebPartToWikiPage(url, wpEntity, (int)webpart.Row, (int)webpart.Column, false);
                            }
                        }
                        var allWebParts = web.GetWebParts(url);
                        foreach (var webpart in allWebParts)
                        {
                            parser.AddToken(new WebPartIdToken(web, webpart.WebPart.Title, webpart.Id));
                        }
                    }

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
                        web.Context.ExecuteQuery();
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


        public override bool WillProvision(Web web, ProvisioningTemplate template)
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
