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
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPageContents : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Page Contents"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // This handler only extracts contents and adds them to the Files and Pages collection.
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Extract the Home Page
                web.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);

                var homepageUrl = web.RootFolder.WelcomePage;
                if (string.IsNullOrEmpty(homepageUrl))
                {
                    homepageUrl = "Default.aspx";
                }
                var welcomePageUrl = UrlUtility.Combine(web.ServerRelativeUrl, homepageUrl);

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

                                    var webPartxml = TokenizeWebPartXml(web, web.GetWebPartXml(webPart.Id, welcomePageUrl));

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

                            // Set the homepage
                            if (template.WebSettings == null)
                            {
                                template.WebSettings = new WebSettings();
                            }
                            template.WebSettings.WelcomePage = homepageUrl;


                        }
                        else
                        {
                            if (web.Context.HasMinimalServerLibraryVersion(Constants.MINIMUMZONEIDREQUIREDSERVERVERSION))
                            {
                                // Not a wikipage
                                template = GetFileContents(web, template, welcomePageUrl, creationInfo, scope);
                                if (template.WebSettings == null)
                                {
                                    template.WebSettings = new WebSettings();
                                }
                                template.WebSettings.WelcomePage = homepageUrl;
                            }
                            else
                            {
                                WriteWarning(string.Format("Page content export requires a server version that is newer than the current server. Server version is {0}, minimal required is {1}", web.Context.ServerLibraryVersion, Constants.MINIMUMZONEIDREQUIREDSERVERVERSION), ProvisioningMessageType.Warning);
                                scope.LogWarning("Page content export requires a server version that is newer than the current server. Server version is {0}, minimal required is {1}", web.Context.ServerLibraryVersion, Constants.MINIMUMZONEIDREQUIREDSERVERVERSION);
                            }
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
                        if (web.Context.HasMinimalServerLibraryVersion(Constants.MINIMUMZONEIDREQUIREDSERVERVERSION))
                        {
                            // Page does not belong to a list, extract the file as is
                            template = GetFileContents(web, template, welcomePageUrl, creationInfo, scope);
                            if (template.WebSettings == null)
                            {
                                template.WebSettings = new WebSettings();
                            }
                            template.WebSettings.WelcomePage = homepageUrl;
                        }
                        else
                        {
                            WriteWarning(string.Format("Page content export requires a server version that is newer than the current server. Server version is {0}, minimal required is {1}", web.Context.ServerLibraryVersion, Constants.MINIMUMZONEIDREQUIREDSERVERVERSION), ProvisioningMessageType.Warning);
                            scope.LogWarning("Page content export requires a server version that is newer than the current server. Server version is {0}, minimal required is {1}", web.Context.ServerLibraryVersion, Constants.MINIMUMZONEIDREQUIREDSERVERVERSION);
                        }
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

        private ProvisioningTemplate GetFileContents(Web web, ProvisioningTemplate template, string welcomePageUrl, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope)
        {
            var fullUri = new Uri(UrlUtility.Combine(web.Url, web.RootFolder.WelcomePage));

            var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
            var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

            var webParts = web.GetWebParts(welcomePageUrl);

            var file = web.GetFileByServerRelativeUrl(welcomePageUrl);

            var homeFile = new Model.File()
            {
                Folder = Tokenize(folderPath, web.Url),
                Src = fileName,
                Overwrite = true,
            };

            // Add field values to file

            RetrieveFieldValues(web, file, homeFile);

            // Add WebParts to file
            foreach (var webPart in webParts)
            {
                var webPartxml = TokenizeWebPartXml(web, web.GetWebPartXml(webPart.Id, welcomePageUrl));

                Model.WebPart newWp = new Model.WebPart()
                {
                    Title = webPart.WebPart.Title,
                    Row = (uint)webPart.WebPart.ZoneIndex,
                    Contents = webPartxml
                };
#if !CLIENTSDKV15
                // As long as we've no CSOM library that has the ZoneID we can't use the version check as things don't compile...
                if (web.Context.HasMinimalServerLibraryVersion(Constants.MINIMUMZONEIDREQUIREDSERVERVERSION))
                {
                    newWp.Zone = webPart.ZoneId;
                }
#endif
                homeFile.WebParts.Add(newWp);
            }
            template.Files.Add(homeFile);

            // Persist file using connector
            if (creationInfo.PersistBrandingFiles)
            {
                PersistFile(web, creationInfo, scope, folderPath, fileName);
            }
            return template;
        }

        private Model.File RetrieveFieldValues(Web web, File file, Model.File homeFile)
        {
            var listItem = file.EnsureProperty(f => f.ListItemAllFields);

            var list = listItem.ParentList;

            var fields = list.Fields;
            web.Context.Load(fields, fs => fs.IncludeWithDefaultProperties(f => f.TypeAsString, f => f.InternalName, f => f.Title));
            web.Context.ExecuteQueryRetry();

            var fieldValues = listItem.FieldValues;

            var fieldValuesAsText = listItem.EnsureProperty(li => li.FieldValuesAsText).FieldValues;

            var fieldstoExclude = new[] {
                "ID",
                "GUID",
                "Author",
                "Editor",
                "FileLeafRef",
                "FileRef",
                "File_x0020_Type",
                "Modified_x0020_By",
                "Created_x0020_By",
                "Created",
                "Modified",
                "FileDirRef",
                "Last_x0020_Modified",
                "Created_x0020_Date",
                "File_x0020_Size",
                "FSObjType",
                "IsCheckedoutToLocal",
                "ScopeId",
                "UniqueId",
                "VirusStatus",
                "_Level",
                "_IsCurrentVersion",
                "ItemChildCount",
                "FolderChildCount",
                "SMLastModifiedDate",
                "owshiddenversion",
                "_UIVersion",
                "_UIVersionString",
                "Order",
                "WorkflowVersion",
                "DocConcurrencyNumber",
                "ParentUniqueId",
                "CheckedOutUserId",
                "SyncClientId",
                "CheckedOutTitle",
                "SMTotalSize",
                "SMTotalFileStreamSize",
                "SMTotalFileCount",
                "ParentVersionString",
                "ParentLeafName",
                "SortBehavior",
                "_ModerationStatus"
            };

            foreach (var fieldValue in fieldValues.Where(f => !fieldstoExclude.Contains(f.Key)))
            {
                if (fieldValue.Value != null && !string.IsNullOrEmpty(fieldValue.Value.ToString()))
                {
                    var field = fields.FirstOrDefault(fs => fs.InternalName == fieldValue.Key);

                    string value = string.Empty;
                    if (field.TypeAsString == "URL")
                    {
                        value = Tokenize(fieldValuesAsText[fieldValue.Key], web.Url);
                    }
                    else {
                        value = Tokenize(fieldValue.Value.ToString(), web.Url);
                    }

                    if (fieldValue.Key == "ContentTypeId")
                    {
                        // Replace the content typeid with a token
                        var ct = list.GetContentTypeById(value);
                        if (ct != null)
                        {
                            value = string.Format("{{contenttypeid:{0}}}", ct.Name);
                        }
                    }
                    homeFile.Properties.Add(fieldValue.Key, value);
                }
            }

            return homeFile;
        }

        private void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string folderPath, string fileName)
        {

            if (creationInfo.FileConnector != null)
            {
                SharePointConnector connector = new SharePointConnector(web.Context, web.Url, "dummy");

                Uri u = new Uri(web.Url);
                if (folderPath.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    folderPath = folderPath.Replace(u.PathAndQuery, "");
                }

                using (Stream s = connector.GetFileStream(fileName, folderPath))
                {
                    if (s != null)
                    {
                        creationInfo.FileConnector.SaveFileStream(fileName, s);
                    }
                }
            }
            else
            {
                WriteWarning("No connector present to persist homepage.", ProvisioningMessageType.Error);
                scope.LogError("No connector present to persist homepage");
            }

        }

        private string TokenizeWebPartXml(Web web, string xml)
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
            xml = xml.Replace("<![CDATA[", "{cdatastart}");
            xml = xml.Replace("]]>", "{cdataend}");
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
                _willProvision = false;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = web.Context.Credentials != null ? true : false;
            }
            return _willExtract.Value;
        }

    }
}
