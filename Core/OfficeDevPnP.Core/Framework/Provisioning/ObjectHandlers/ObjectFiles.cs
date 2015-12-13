using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client.WebParts;
using System.Xml.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.IO;
using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectFiles : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Files"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);

                foreach (var file in template.Files)
                {
                    var folderName = parser.ParseString(file.Folder);

                    if (folderName.ToLower().StartsWith((web.ServerRelativeUrl.ToLower())))
                    {
                        folderName = Tokenize(folderName.Substring(web.ServerRelativeUrl.Length), web.Url);
                    }

                    var folder = web.EnsureFolderPath(folderName);

                    File targetFile = null;

                    var checkedOut = false;

                    targetFile = folder.GetFile(template.Connector.GetFilenamePart(file.Src));

                    if (targetFile != null)
                    {
                        if (file.Overwrite)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Uploading_and_overwriting_existing_file__0_, file.Src);
                            checkedOut = CheckOutIfNeeded(web, targetFile);

                            using (var stream = template.Connector.GetFileStream(file.Src))
                            {
                                targetFile = folder.UploadFile(template.Connector.GetFilenamePart(file.Src), stream, file.Overwrite);
                            }
                        }
                        else
                        {
                            checkedOut = CheckOutIfNeeded(web, targetFile);
                        }
                    }
                    else
                    {
                        using (var stream = template.Connector.GetFileStream(file.Src))
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Uploading_file__0_, file.Src);
                            targetFile = folder.UploadFile(template.Connector.GetFilenamePart(file.Src), stream, file.Overwrite);
                        }

                        checkedOut = CheckOutIfNeeded(web, targetFile);
                    }

                    if (targetFile != null)
                    {
                        if (file.Properties != null && file.Properties.Any())
                        {
                            Dictionary<string, string> transformedProperties = file.Properties.ToDictionary(property => property.Key, property => parser.ParseString(property.Value));
                            targetFile.SetFileProperties(transformedProperties, false); // if needed, the file is already checked out
                        }

                        if (file.WebParts != null && file.WebParts.Any())
                        {
                            targetFile.EnsureProperties(f => f.ServerRelativeUrl);

                            var existingWebParts = web.GetWebParts(targetFile.ServerRelativeUrl);
                            foreach (var webpart in file.WebParts)
                            {
                                // check if the webpart is already set on the page
                                if (existingWebParts.FirstOrDefault(w => w.WebPart.Title == webpart.Title) == null)
                                {
                                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Adding_webpart___0___to_page, webpart.Title);
                                    var wpEntity = new WebPartEntity();
                                    wpEntity.WebPartTitle = webpart.Title;
                                    wpEntity.WebPartXml = parser.ParseString(webpart.Contents).Trim(new[] { '\n', ' ' });
                                    wpEntity.WebPartZone = webpart.Zone;
                                    wpEntity.WebPartIndex = (int)webpart.Order;
                                    web.AddWebPartToWebPartPage(targetFile.ServerRelativeUrl, wpEntity);
                                }
                            }
                        }

                        if (checkedOut)
                        {
                            targetFile.CheckIn("", CheckinType.MajorCheckIn);
                            web.Context.ExecuteQueryRetry();
                        }

                        // Don't set security when nothing is defined. This otherwise breaks on files set outside of a list
                        if (file.Security != null &&
                            (file.Security.ClearSubscopes == true || file.Security.CopyRoleAssignments == true || file.Security.RoleAssignments.Count > 0))
                        {
                            targetFile.ListItemAllFields.SetSecurity(parser, file.Security);
                        }
                    }

                }
            }
            return parser;
        }

        private static bool CheckOutIfNeeded(Web web, File targetFile)
        {
            var checkedOut = false;
            try
            {
                web.Context.Load(targetFile, f => f.CheckOutType, f => f.ListItemAllFields.ParentList.ForceCheckout);
                web.Context.ExecuteQueryRetry();

                if (targetFile.ListItemAllFields.ServerObjectIsNull.HasValue && !targetFile.ListItemAllFields.ServerObjectIsNull.Value)
                {
                    if (targetFile.CheckOutType == CheckOutType.None)
                    {
                        targetFile.CheckOut();
                    }
                    checkedOut = true;
                }
            }
            catch (ServerException ex)
            {
                // Handling the exception stating the "The object specified does not belong to a list."
                if (ex.ServerErrorCode != -2146232832)
                {
                    throw;
                }
            }
            return checkedOut;
        }


        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Extract the Home PAge
                web.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);

                var welcomePageUrl = UrlUtility.Combine(web.ServerRelativeUrl, web.RootFolder.WelcomePage);

                var file = web.GetFileByServerRelativeUrl(welcomePageUrl);
                try
                {
                    var listItem = file.EnsureProperty(f => f.ListItemAllFields);
                    if (listItem != null)
                    {
                        if (listItem.FieldValues.ContainsKey("WikiField"))
                        {
                            // Wiki page
                            var fullUri = new Uri(UrlUtility.Combine(web.Url, web.RootFolder.WelcomePage));

                            var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
                            var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

                            var homeFile = web.GetFileByServerRelativeUrl(welcomePageUrl);

                            LimitedWebPartManager limitedWPManager =
                                homeFile.GetLimitedWebPartManager(PersonalizationScope.Shared);

                            web.Context.Load(limitedWPManager);

                            var webParts = web.GetWebParts(welcomePageUrl);

                            var page = new Page()
                            {
                                Layout = WikiPageLayout.Custom,
                                Overwrite = true,
                                Url = Tokenize(fullUri.PathAndQuery, web.Url),
                                WelcomePage = true
                            };
                            var pageContents = listItem.FieldValues["WikiField"].ToString();

                            Regex regexClientIds = new Regex(@"id=\""div_(?<ControlId>(\w|\-)+)");
                            if (regexClientIds.IsMatch(pageContents))
                            {
                                foreach (Match webPartMatch in regexClientIds.Matches(pageContents))
                                {
                                    String serverSideControlId = webPartMatch.Groups["ControlId"].Value;

                                    WebPartDefinition webPart = limitedWPManager.WebParts.GetByControlId(String.Format("g_{0}",
                                        serverSideControlId.Replace("-", "_")));
                                    web.Context.Load(webPart,
                                        wp => wp.Id,
                                        wp => wp.WebPart.Title,
                                        wp => wp.WebPart.ZoneIndex
                                        );
                                    web.Context.ExecuteQueryRetry();

                                    var webPartxml = Tokenize(web, web.GetWebPartXml(webPart.Id, welcomePageUrl));

                                    page.WebParts.Add(new Model.WebPart()
                                    {
                                        Title = webPart.WebPart.Title,
                                        Contents = webPartxml,
                                        Order = (uint)webPart.WebPart.ZoneIndex,
                                        Row = 1, // By default we will create a onecolumn layout, add the webpart to it, and later replace the wikifield on the page to position the webparts correctly.
                                        Column = 1 // By default we will create a onecolumn layout, add the webpart to it, and later replace the wikifield on the page to position the webparts correctly.
                                    });

                                    pageContents = Regex.Replace(pageContents, serverSideControlId, string.Format("{{webpartid:{0}}}", webPart.WebPart.Title), RegexOptions.IgnoreCase);
                                }
                            }

                            page.Fields.Add("WikiField", pageContents);
                            template.Pages.Add(page);


                        }
                        else
                        {
                            // Not a wikipage
                            template = GetFileContents(web, template, welcomePageUrl);
                        }
                    }
                }
                catch (ServerException ex)
                {
                    if (ex.ServerErrorCode != -2146232832)
                    {
                        throw;
                    }
                    else
                    {
                        // Page does not belong to a list, extract the file as is
                        template = GetFileContents(web, template, welcomePageUrl);
                    }
                }

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate GetFileContents(Web web, ProvisioningTemplate template, string welcomePageUrl)
        {
            var fullUri = new Uri(UrlUtility.Combine(web.Url, web.RootFolder.WelcomePage));

            var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
            var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

            var webParts = web.GetWebParts(welcomePageUrl);

            var homeFile = new Model.File()
            {
                Folder = Tokenize(folderPath, web.Url),
                Src = fileName,
                Overwrite = true,
            };

            foreach (var webPart in webParts)
            {
                var webPartxml = Tokenize(web, web.GetWebPartXml(webPart.Id, welcomePageUrl));

                homeFile.WebParts.Add(new Model.WebPart()
                {
                    Title = webPart.WebPart.Title,
                    Row = (uint)webPart.WebPart.ZoneIndex,
                    Zone = webPart.ZoneId,
                    Contents = webPartxml
                });
            }
            template.Files.Add(homeFile);

            return template;
        }

        private string Tokenize(Web web, string xml)
        {
            var lists = web.Lists;
            web.Context.Load(web, w => w.ServerRelativeUrl, w => w.Id);
            web.Context.Load(lists, ls => ls.Include(l => l.Id, l => l.Title));
            web.Context.ExecuteQueryRetry();

            foreach (var list in lists)
            {
                xml = Regex.Replace(xml, list.Id.ToString(), string.Format("{{listid:{0}}}", list.Title), RegexOptions.IgnoreCase);
            }
            xml = Regex.Replace(xml, web.Id.ToString(), "{siteid}", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, web.ServerRelativeUrl, "{site}", RegexOptions.IgnoreCase);

            return xml;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {

            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Files.Any();
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
