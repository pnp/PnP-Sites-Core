using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PnPFolder = OfficeDevPnP.Core.Framework.Provisioning.Model.Folder;
using PnPSiteSecurity = OfficeDevPnP.Core.Framework.Provisioning.Model.SiteSecurity;
using PnPFile = OfficeDevPnP.Core.Framework.Provisioning.Model.File;
using PnPFileLevel = OfficeDevPnP.Core.Framework.Provisioning.Model.FileLevel;
using PnPRoleAssignment = OfficeDevPnP.Core.Framework.Provisioning.Model.RoleAssignment;
using PnPDataRow = OfficeDevPnP.Core.Framework.Provisioning.Model.DataRow;
using PnPField = OfficeDevPnP.Core.Framework.Provisioning.Model.Field;
using PnPView = OfficeDevPnP.Core.Framework.Provisioning.Model.View;
using SPFolder = Microsoft.SharePoint.Client.Folder;
using SPField = Microsoft.SharePoint.Client.Field;
using SPFile = Microsoft.SharePoint.Client.File;
using SPList = Microsoft.SharePoint.Client.List;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectListInstanceDataRows : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "List instances Data Rows"; }
        }

        public override string InternalName => "ListInstanceDataRows";
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (!template.Lists.Any()) return parser;

                web.EnsureProperties(w => w.ServerRelativeUrl);

                web.Context.Load(web.Lists, lc => lc.IncludeWithDefaultProperties(l => l.RootFolder.ServerRelativeUrl));
                web.Context.ExecuteQueryRetry();

                #region DataRows

                foreach (var listInstance in template.Lists)
                {
                    if (listInstance.DataRows != null && listInstance.DataRows.Any())
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Processing_data_rows_for__0_, listInstance.Title);
                        // Retrieve the target list
                        var list = web.Lists.GetByTitle(parser.ParseString(listInstance.Title));
                        web.Context.Load(list);

                        // Retrieve the fields' types from the list
                        Microsoft.SharePoint.Client.FieldCollection fields = list.Fields;
                        web.Context.Load(fields, fs => fs.Include(f => f.InternalName, f => f.FieldTypeKind, f => f.TypeAsString, f => f.ReadOnlyField, f => f.Title));
                        web.Context.ExecuteQueryRetry();

                        var keyColumnType = "Text";
                        var parsedKeyColumn = parser.ParseString(listInstance.DataRows.KeyColumn);
                        if (!string.IsNullOrEmpty(parsedKeyColumn))
                        {
                            var keyColumn = fields.FirstOrDefault(f => f.InternalName.Equals(parsedKeyColumn, StringComparison.InvariantCultureIgnoreCase));
                            if (keyColumn != null)
                            {
                                switch (keyColumn.FieldTypeKind)
                                {
                                    case FieldType.User:
                                    case FieldType.Lookup:
                                        keyColumnType = "Lookup";
                                        break;

                                    case FieldType.URL:
                                        keyColumnType = "Url";
                                        break;

                                    case FieldType.DateTime:
                                        keyColumnType = "DateTime";
                                        break;

                                    case FieldType.Number:
                                    case FieldType.Counter:
                                        keyColumnType = "Number";
                                        break;
                                }
                            }
                        }

                        foreach (var dataRow in listInstance.DataRows)
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_list_item__0_, listInstance.DataRows.IndexOf(dataRow) + 1);

                                bool processItem = true;
                                ListItem listitem = null;

                                if (!string.IsNullOrEmpty(listInstance.DataRows.KeyColumn))
                                {
                                    // Get value from key column
                                    var dataRowValues = dataRow.Values.Where(v => v.Key == listInstance.DataRows.KeyColumn).ToList();

                                    // if it is empty, skip the check
                                    if (dataRowValues.Any())
                                    {
                                        var query = $@"<View><Query><Where><Eq><FieldRef Name=""{parsedKeyColumn}""/><Value Type=""{keyColumnType}"">{parser.ParseString(dataRowValues.FirstOrDefault().Value)}</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>";
                                        var camlQuery = new CamlQuery()
                                        {
                                            ViewXml = query
                                        };
                                        var existingItems = list.GetItems(camlQuery);
                                        list.Context.Load(existingItems);
                                        list.Context.ExecuteQueryRetry();
                                        if (existingItems.Count > 0)
                                        {
                                            if (listInstance.DataRows.UpdateBehavior == UpdateBehavior.Skip)
                                            {
                                                processItem = false;
                                            }
                                            else
                                            {
                                                listitem = existingItems[0];
                                                processItem = true;
                                            }
                                        }
                                    }
                                }

                                if (processItem)
                                {
                                    if (listitem == null)
                                    {
                                        var listitemCI = new ListItemCreationInformation();
                                        listitem = list.AddItem(listitemCI);
                                    }

                                    ListItemUtilities.UpdateListItem(listitem, parser, dataRow.Values, ListItemUtilities.ListItemUpdateType.UpdateOverwriteVersion);

                                    if (dataRow.Security != null && (dataRow.Security.ClearSubscopes || dataRow.Security.CopyRoleAssignments || dataRow.Security.RoleAssignments.Count > 0))
                                    {
                                        listitem.SetSecurity(parser, dataRow.Security);
                                    }

                                    if (dataRow.Attachments != null && dataRow.Attachments.Count > 0)
                                    {
                                        foreach (var attachment in dataRow.Attachments)
                                        {
                                            attachment.Name = parser.ParseString(attachment.Name);
                                            attachment.Src = parser.ParseString(attachment.Src);
                                            var overwrite = attachment.Overwrite;
                                            listitem.EnsureProperty(l => l.AttachmentFiles);

                                            Attachment existingItem = null;
                                            if (listitem.AttachmentFiles.Count > 0)
                                            {
                                                existingItem = listitem.AttachmentFiles.FirstOrDefault(a => a.FileName.Equals(attachment.Name, StringComparison.OrdinalIgnoreCase));
                                            }
                                            if (existingItem != null)
                                            {
                                                if (overwrite)
                                                {
                                                    existingItem.DeleteObject();
                                                    web.Context.ExecuteQueryRetry();
                                                    AddAttachment(template, listitem, attachment);
                                                }
                                            }
                                            else
                                            {
                                                AddAttachment(template, listitem, attachment);
                                            }
                                        }
                                    }
                                }
                            }
                            catch (ServerException ex)
                            {
                                if (ex.ServerErrorTypeName.Equals("Microsoft.SharePoint.SPDuplicateValuesFoundException", StringComparison.InvariantCultureIgnoreCase)
                                    && applyingInformation.IgnoreDuplicateDataRowErrors)
                                {
                                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_duplicate);
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_failed___0_____1_, ex.Message, ex.StackTrace);
                                throw;
                            }
                        }
                    }
                }


                #endregion DataRows
            }


            return parser;
        }


        private void AddAttachment(ProvisioningTemplate template, ListItem listitem, Model.SharePoint.InformationArchitecture.DataRowAttachment attachment)
        {
#if !SP2013 && !SP2016
            listitem.AttachmentFiles.AddUsingPath(ResourcePath.FromDecodedUrl(attachment.Name), FileUtilities.GetFileStream(template, attachment.Src));
#else
            var attachmentCI = new AttachmentCreationInformation()
            {
                ContentStream = FileUtilities.GetFileStream(template, attachment.Src),
                FileName = attachment.Name
            };
            listitem.AttachmentFiles.Add(attachmentCI);
#endif
            listitem.Context.ExecuteQueryRetry();
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (creationInfo.IncludeAllListContent)
            {
                //Export all List and Library Context exclude ClientSidePages and derived ContentType Files since handled in ObjectClientSidePageContent
                using (var scope = new PnPMonitoredScope(this.Name))
                {
                    int total = template.Lists.Count;
                    var listCount = 0;

                    //all Libs but not SitePage
                    foreach (ListInstance listInstance in template.Lists.Where(t => t.TemplateType != 119))
                    {
                        creationInfo.MessagesDelegate?.Invoke($"ListData|{listInstance.Url}|{listCount}|{total}", ProvisioningMessageType.Progress);

                        SPList myList = web.GetListByUrl(listInstance.Url);
                        web.Context.Load(myList, l => l.BaseType, l => l.Id);
                        var fields = myList.Fields;
                        web.Context.Load(fields, fs => fs.Include(f => f.TypeAsString, f => f.InternalName, f => f.ReadOnlyField, f => f.FieldTypeKind, f => f.Title));
                        web.Context.ExecuteQueryRetry();

                        ListItemCollection listItems = myList.GetItems(new CamlQuery() { ViewXml = "<View Scope=\"RecursiveAll\"><Query><OrderBy><FieldRef Name='ID' Ascending='True'></FieldRef></OrderBy></Query></View>" });
                        web.Context.Load(listItems, lc => lc.Include(li => li.Id, li => li.FieldValuesAsText, li => li.FileSystemObjectType, li => li["Title"]));
                        web.Context.ExecuteQueryRetry();

                        if (myList.BaseType == BaseType.DocumentLibrary)
                        {
                            foreach (var spItem in listItems)
                            {
                                switch (spItem.FileSystemObjectType)
                                {
                                    case FileSystemObjectType.File:
                                        {
                                            //PnP:File
                                            ProcessDocumentRow(web, spItem, listInstance, template, scope, creationInfo.FilesToIgnore);
                                            break;
                                        }
                                    case FileSystemObjectType.Folder:
                                        {
                                            //PnP:Folder
                                            ProcessFolderRow(web, spItem, listInstance, template, scope);
                                            break;
                                        }
                                    default:
                                        {
                                            //PnP:DataRow
                                            ProcessDataRow(web, spItem, listInstance, template, scope);
                                            break;
                                        }
                                }
                            }
                            //Check if we have Templates in Forms Folder
                            ProcessFormsFolder(web, myList, listInstance, template, scope);
                        }
                        else
                        {
                            foreach (var spItem in listItems)
                            {
                                //PnP:DataRow
                                ProcessDataRow(web, spItem, listInstance, template, scope);
                            }
                        }

                        //fix: Columns Added on List Level are Added to listInstance.FieldRefs
                        if (listInstance.Fields != null && listInstance.Fields.Any())
                        {
                            foreach (PnPField siteField in listInstance.Fields)
                            {
                                Guid fieldId = Guid.Empty;
                                XElement fieldSchema = XElement.Parse(siteField.SchemaXml);
                                if (Guid.TryParse((string)fieldSchema.Attribute("ID"), out fieldId))
                                {
                                    if (!listInstance.FieldRefs.Any(f => f.Id == fieldId))
                                    {
                                        var spField = myList.GetFieldById(fieldId);
                                        if (spField != null)
                                        {
                                            try
                                            {
                                                web.Context.Load(spField, f => f.SchemaXml);
                                                web.Context.ExecuteQuery();
                                                XElement fSchema = XElement.Parse(spField.SchemaXml);
                                                string fieldName = (string)fSchema.Attribute("Name");
                                                string fieldDisplayName = (string)fSchema.Attribute("DisplayName");
                                                listInstance.FieldRefs.Add(new FieldRef(fieldName) { Id = fieldId, DisplayName = fieldDisplayName });
                                            }
                                            catch (Exception ex)
                                            {

                                            }
                                        }
                                    }
                                }
                            }
                        }

                        listCount++;
                    }

                    //Handle SitePages-Library Content
                    //Some Content is already taken care of by ObjectPageContents and ObjectClientSidePageContents
                    //ObjectPageContents does not export security
                    //ObjectClientSidePageContents does not care about Security and additional Fields from CustomContentType
                    if (template.Lists.Any(t => t.TemplateType != 119) && creationInfo.HandlersToProcess.HasFlag(Handlers.PageContents))
                    {
                        //todo: handle all whats not .aspx or taking care of everything not yet extracted i.e. additional Fields, Security, ContentType
                        var sitePagesLib = template.Lists.FirstOrDefault(t => t.TemplateType != 119);

                    }
                }
            }
            return template;
        }

        /// <summary>
        /// Extract File in Document Library
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listItem"></param>
        /// <param name="listInstance"></param>
        /// <param name="template"></param>
        /// <param name="scope"></param>
        /// <param name="excludeFileNames"></param>
        private void ProcessDocumentRow(Web web, ListItem listItem, ListInstance listInstance, ProvisioningTemplate template, PnPMonitoredScope scope, List<String>excludeFileNames)
        {
            web.EnsureProperties(w => w.Url, w => w.Title);
            SPFile myFile = listItem.File;
            web.Context.Load(myFile,
                f => f.Name,
                f => f.ServerRelativeUrl,
                f => f.Level,
                f => f.UniqueId);
            web.Context.ExecuteQueryRetry();

            // If we got here it's a file, let's grab the file's path and name
            var baseUri = new Uri(web.Url);
            var fullUri = new Uri(baseUri, myFile.ServerRelativeUrl);
            var folderPath = System.Web.HttpUtility.UrlDecode(fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
            var fileName = System.Web.HttpUtility.UrlDecode(fullUri.Segments[fullUri.Segments.Count() - 1]);

            var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

            //ignore certain files
            if (excludeFileNames != null)
            {
                if (excludeFileNames.Any(f => f.Equals($"{templateFolderPath}/{fileName}", StringComparison.InvariantCultureIgnoreCase)))
                    return;
            }

            // Avoid duplicate file entries
            PnPFile newFile = null;
            bool addFile = false;

            newFile = template.Files.FirstOrDefault(f => f.Src.Equals($"{templateFolderPath}/{fileName}", StringComparison.CurrentCultureIgnoreCase));
            if (newFile == null)
            {

                newFile = new PnPFile()
                {
                    Folder = templateFolderPath,
                    Src = $"{templateFolderPath}/{fileName}",
                    TargetFileName = myFile.Name,
                    Overwrite = true,
                    Level = (PnPFileLevel)Enum.Parse(typeof(PnPFileLevel), myFile.Level.ToString())
                };
                addFile = true;
            }

            ExtractFileSettings(web, myFile.UniqueId, ref newFile, template.Security, scope, FileFieldsToExclude);

            if (addFile)
            {
                SPFile file = listItem.File;
                web.Context.Load(file);
                web.Context.ExecuteQueryRetry();
                var spFileStream = file.OpenBinaryStream();
                web.Context.ExecuteQueryRetry();

                template.Connector.SaveFileStream(file.Name, templateFolderPath, spFileStream.Value);

                template.Files.Add(newFile);
            }
        }

        /// <summary>
        /// Extract FieldValues for File
        /// </summary>
        /// <param name="web"></param>
        /// <param name="fileUniqueId"></param>
        /// <param name="pnpFile"></param>
        /// <param name="siteSecurity"></param>
        /// <param name="scope"></param>
        /// <param name="fieldsToExclude"></param>
        private void ExtractFileSettings(Web web, Guid fileUniqueId, ref PnPFile pnpFile, PnPSiteSecurity siteSecurity, PnPMonitoredScope scope, string[] fieldsToExclude = null)
        {
            try
            {
                var file = web.GetFileById(fileUniqueId);
                web.Context.Load(file,
                    f => f.Level,
                    f => f.ServerRelativeUrl,
                    f => f.Properties,
                    f => f.ListItemAllFields,
                    f => f.ListItemAllFields.RoleAssignments,
                    f => f.ListItemAllFields.RoleAssignments.Include(r => r.Member, r => r.RoleDefinitionBindings),
                    f => f.ListItemAllFields.HasUniqueRoleAssignments,
                    f => f.ListItemAllFields.ParentList,
                    f => f.ListItemAllFields.ContentType.StringId);
                web.Context.Load(web,
                    w => w.AssociatedOwnerGroup,
                    w => w.AssociatedMemberGroup,
                    w => w.AssociatedVisitorGroup,
                    w => w.Title,
                    w => w.RoleDefinitions.Include(r => r.RoleTypeKind, r => r.Name),
                    w => w.ContentTypes.Include(c => c.Id, c => c.Name, c => c.StringId));

                web.Context.ExecuteQueryRetry();

                ////export PnPFile Properties
                //if (file.Properties.FieldValues.Any())
                //{
                //    foreach (var propKey in file.Properties.FieldValues.Keys.Where(k => !k.StartsWith("vti_") && !k.StartsWith("docset_")))
                //    {
                //        //nowhere to store in schema, currently no need
                //    }
                //}

                //export PnPFile FieldValues
                if (file.ListItemAllFields.FieldValues.Any())
                {
                    var list = file.ListItemAllFields.ParentList;

                    var fields = list.Fields;
                    web.Context.Load(fields, fs => fs.IncludeWithDefaultProperties(f => f.TypeAsString, f => f.InternalName, f => f.Title));
                    web.Context.ExecuteQueryRetry();

                    var fieldValues = file.ListItemAllFields.FieldValues;

                    var fieldValuesAsText = file.ListItemAllFields.EnsureProperty(li => li.FieldValuesAsText).FieldValues;

                    if (fieldsToExclude == null)
                    {
                        fieldsToExclude = new string[] { };
                    }

                    #region //**** get correct Content Type
                    string ctId = string.Empty;
                    foreach (var ct in web.ContentTypes.OrderByDescending(c => c.StringId.Length))
                    {
                        if (file.ListItemAllFields.ContentType.StringId.StartsWith(ct.StringId))
                        {
                            pnpFile.Properties.Add("ContentTypeId", ct.StringId);
                            break;
                        }
                    }
                    #endregion //**** get correct Content Type

                    foreach (var fieldValue in fieldValues.Where(f => !fieldsToExclude.Contains(f.Key)))
                    {
                        if (fieldValue.Value != null && !string.IsNullOrEmpty(fieldValue.Value.ToString()))
                        {
                            var field = fields.FirstOrDefault(fs => fs.InternalName == fieldValue.Key);
                            string value = string.Empty;

                            //ignore read only fields
                            if (!field.ReadOnlyField || WriteableReadOnlyField.Contains(field.InternalName.ToLower()))
                            {
                                value = TokenizeValue(web, field, fieldValue, fieldValuesAsText);

                                if (fieldValue.Key == "ContentTypeId")
                                {
                                    value = null; //it's already in Properties - we can ignore here
                                }
                            }

                            // We process real values only
                            if (value != null && !String.IsNullOrEmpty(value) && value != "[]")
                            {
                                pnpFile.Properties.Add(fieldValue.Key, value);
                            }
                        }
                    }

                    //get PnPFile Permissions
                    if (file.ListItemAllFields.HasUniqueRoleAssignments) // && siteSecurity != null)
                    {
                        GetObjectSecurity(web, file.ListItemAllFields.RoleAssignments, pnpFile.Security);
                    }
                }
            }
            catch (Exception ex)
            {
                scope.LogError(ex, "Extract of File with uniqueId {0} failed", fileUniqueId);
            }
        }

        /// <summary>
        /// Extract File Referenced in NewDocumentTemplates
        /// </summary>
        /// <param name="web"></param>
        /// <param name="spList"></param>
        /// <param name="listInstance"></param>
        /// <param name="template"></param>
        /// <param name="scope"></param>
        private void ProcessFormsFolder(Web web, SPList spList, ListInstance listInstance, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            SPFolder formsFolder = null;
            try
            {
                web.EnsureProperties(w => w.Url);
                spList.EnsureProperties(l => l.RootFolder.ServerRelativeUrl);
                formsFolder = web.GetFolderByServerRelativeUrl(spList.RootFolder.ServerRelativeUrl + "/Forms");
                web.Context.ExecuteQueryRetry();
            }
            catch(Exception ex)
            {
                formsFolder = null;
            }
            if (formsFolder != null)
            {
                var baseUri = new Uri(web.Url);

                foreach (PnPView instanceView in listInstance.Views)
                {
                    if (instanceView.SchemaXml.Contains("NewDocumentTemplates"))
                    {
                        var viewSchema = XDocument.Parse(instanceView.SchemaXml);
                        var templateElement = viewSchema.Root.Elements().FirstOrDefault(element => element.Name.LocalName == "NewDocumentTemplates");
                        if(templateElement!=null)
                        {
                            var NewDocumentTemplates = Newtonsoft.Json.Linq.JArray.Parse(templateElement.Value);
                            foreach(var jTok in NewDocumentTemplates)
                            {
                                var jObj = jTok as Newtonsoft.Json.Linq.JObject;
                                if (jObj != null)
                                {
                                    var contentTypeId = jObj.Children<Newtonsoft.Json.Linq.JProperty>().FirstOrDefault(p => p.Name == "contentTypeId");
                                    var url = jObj.Children<Newtonsoft.Json.Linq.JProperty>().FirstOrDefault(p => p.Name == "url");

                                    if(contentTypeId!=null && url!=null)
                                    {
                                        string templateDocUrl = url.Value.ToString();
                                        var fullUri = new Uri(baseUri, templateDocUrl.Replace("{site}", baseUri.AbsolutePath.TrimEnd(new char[] { '/' })));
                                        var folderPath = System.Web.HttpUtility.UrlDecode(fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
                                        var fileName = System.Web.HttpUtility.UrlDecode(fullUri.Segments[fullUri.Segments.Count() - 1]);

                                        var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

                                        SPFile myFile = web.GetFileByUrl($"{templateFolderPath}/{fileName}");
                                        web.Context.Load(myFile);
                                        var stream = myFile.OpenBinaryStream();
                                        web.Context.ExecuteQueryRetry();

                                        template.Connector.SaveFileStream(myFile.Name, templateFolderPath, stream.Value);

                                        PnPFile newFile = new PnPFile()
                                        {
                                            Folder = templateFolderPath,
                                            Src = $"{templateFolderPath}/{fileName}",
                                            TargetFileName = myFile.Name,
                                            Overwrite = true,
                                            Level = (PnPFileLevel)Enum.Parse(typeof(PnPFileLevel), myFile.Level.ToString())
                                        };

                                        template.Files.Add(newFile);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Extract Folder in DocumentLibrary
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listItem"></param>
        /// <param name="listInstance"></param>
        /// <param name="template"></param>
        /// <param name="scope"></param>
        private void ProcessFolderRow(Web web, ListItem listItem, ListInstance listInstance, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            listItem.EnsureProperties(it => it.ParentList.RootFolder.ServerRelativeUrl);
            string serverRelativeListUrl = listItem.ParentList.RootFolder.ServerRelativeUrl;
            string folderPath = listItem.FieldValuesAsText["FileRef"].Substring(serverRelativeListUrl.Length).TrimStart(new char[] { '/' });

            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                listItem.EnsureProperties(it => it.Folder.UniqueId);
                string[] folderSegments = folderPath.Split('/');
                PnPFolder pnpFolder = null;
                for (int i = 0; i < folderSegments.Length; i++)
                {
                    if (i == 0)
                    {
                        pnpFolder = listInstance.Folders.FirstOrDefault(f => f.Name.Equals(folderSegments[i], StringComparison.CurrentCultureIgnoreCase));
                        if (pnpFolder == null)
                        {
                            string pathToCurrentFolder = string.Format("{0}/{1}", serverRelativeListUrl, string.Join("/", folderSegments.Take(i + 1)));
                            pnpFolder = ExtractFolderSettings(web, pathToCurrentFolder, template.Security, scope, FolderFieldsToExclude);
                            listInstance.Folders.Add(pnpFolder);
                        }
                    }
                    else
                    {
                        var childFolder = pnpFolder.Folders.FirstOrDefault(f => f.Name.Equals(folderSegments[i], StringComparison.CurrentCultureIgnoreCase));
                        if (childFolder == null)
                        {
                            string pathToCurrentFolder = string.Format("{0}/{1}", serverRelativeListUrl, string.Join("/", folderSegments.Take(i + 1)));
                            childFolder = ExtractFolderSettings(web, pathToCurrentFolder, template.Security, scope, FolderFieldsToExclude);
                            pnpFolder.Folders.Add(childFolder);
                        }
                        pnpFolder = childFolder;
                    }
                }
            }
        }

        /// <summary>
        /// Extract List Data (with Attachments as soon as Schema support is there)
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listItem"></param>
        /// <param name="listInstance"></param>
        /// <param name="template"></param>
        /// <param name="scope"></param>
        private void ProcessDataRow(Web web, ListItem listItem, ListInstance listInstance, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            try
            {
                SPList myList = web.GetListByUrl(listInstance.Url);
                ListItemCollection listItems = myList.GetItems(new CamlQuery() { ViewXml = "<View Scope=\"RecursiveAll\"><Query><OrderBy><FieldRef Name='ID' Ascending='True'></FieldRef></OrderBy></Query></View>" });
                web.Context.Load(listItems, co => co.Include(cc => cc.Id, cc => cc["Title"]));
                web.Context.ExecuteQueryRetry();

                foreach (var spItem in listItems)
                {
                    //PnPDataRow
                    var pnpDataRow = ExtractListItemSettings(web, myList, spItem.Id, listInstance, scope, ExcludeCustomListFields);
                    if (pnpDataRow != null)
                        listInstance.DataRows.Add(pnpDataRow);
                }
            }
            catch (Exception ex)
            {
                scope.LogError(ex, "Extract of ListData from {0} failed", listInstance.Url);
            }
        }

        private PnPDataRow ExtractListItemSettings(Web web, SPList spList, int ListItemId, ListInstance listInstance, PnPMonitoredScope scope, string[] fieldsToExclude = null)
        {
            try
            {
                web.EnsureProperties(w => w.Url, w => w.ContentTypes.Include(c => c.Id, c => c.Name, c => c.StringId));

                var fields = spList.Fields;
                web.Context.Load(fields, fs => fs.Include(f => f.InternalName, f => f.FieldTypeKind, f => f.TypeAsString, f => f.ReadOnlyField, f => f.Title));
                web.Context.ExecuteQueryRetry();

                var listItem = spList.GetItemById(ListItemId);
                web.Context.Load(listItem);
                web.Context.Load(listItem, li => li.AttachmentFiles.Include(a => a.FileName, a => a.ServerRelativeUrl), li => li.ContentType.StringId);
                web.Context.ExecuteQueryRetry();

                var fieldValues = listItem.FieldValues;
                var fieldValuesAsText = listItem.EnsureProperty(li => li.FieldValuesAsText).FieldValues;

                if (fieldsToExclude == null)
                {
                    fieldsToExclude = new string[] { };
                }

                Dictionary<string, string> dataRow = new Dictionary<string, string>();
                #region //**** get correct Content Type
                string ctId = string.Empty;
                foreach (var ct in web.ContentTypes.OrderByDescending(c => c.StringId.Length))
                {
                    if (listItem.ContentType.StringId.StartsWith(ct.StringId))
                    {
                        dataRow.Add("ContentTypeId", ct.StringId);
                        break;
                    }
                }
                #endregion //**** get correct Content Type

                foreach (var fieldValue in listItem.FieldValues.Where(f => !fieldsToExclude.Contains(f.Key)))
                {
                    if (fieldValue.Value != null && !string.IsNullOrEmpty(fieldValue.Value.ToString()))
                    {
                        var field = fields.FirstOrDefault(fs => fs.InternalName == fieldValue.Key);
                        string value = string.Empty;

                        //ignore read only fields
                        if (!field.ReadOnlyField || WriteableReadOnlyListField.Contains(field.InternalName.ToLower()))
                        {
                            value = TokenizeValue(web, field, fieldValue, fieldValuesAsText);
                        }

                        if (fieldValue.Key.Equals("ContentTypeId", StringComparison.CurrentCultureIgnoreCase))
                        {
                            value = null; //ignore here since already in dataRow
                        }

                        // We process real values only
                        if (!string.IsNullOrWhiteSpace(value) && value != "[]")
                        {
                            dataRow.Add(fieldValue.Key, value);
                        }
                    }
                }

                var pnpDataRow= new PnPDataRow(dataRow);

                if (listItem.AttachmentFiles.Count > 0)
                {
                    Dictionary<string, string> AttachmentFiles = new Dictionary<string, string>();
                    foreach (Attachment file in listItem.AttachmentFiles)
                    {
                        AttachmentFiles.Add(file.FileName, file.ServerRelativeUrl);
                    }
                    SaveAttachmentsToConnector(web, listInstance, listItem.Id, AttachmentFiles, pnpDataRow, scope);
                }


                return pnpDataRow;
            }
            catch (Exception ex)
            {
                scope.LogError(ex, "Extract of ListItemId {0} failed", ListItemId);
            }
            return null;
        }

        /// <summary>
        /// Save List Attachments
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listInstance"></param>
        /// <param name="ListItemId"></param>
        /// <param name="AttachmentFiles"></param>
        /// <param name="dataRow"></param>
        /// <param name="scope"></param>
        private void SaveAttachmentsToConnector(Web web, ListInstance listInstance, int ListItemId, Dictionary<string, string> AttachmentFiles, PnPDataRow dataRow, PnPMonitoredScope scope)
        {
            foreach (var fileName in AttachmentFiles.Keys)
            {
                string destfileFolder = $"{listInstance.Url.Remove(0, 6)}/{ListItemId}/";
                try
                {
                    string serverRelativeUrl = AttachmentFiles[fileName];
                    var spFile = web.GetFileByServerRelativeUrl(serverRelativeUrl);
                    web.Context.Load(spFile);
                    web.Context.ExecuteQueryRetry();
                    var streamX = spFile.OpenBinaryStream();
                    web.Context.ExecuteQueryRetry();
                    listInstance.ParentTemplate.Connector.SaveFileStream(fileName, destfileFolder, streamX.Value);
                    dataRow.Attachments.Add(new Model.SharePoint.InformationArchitecture.DataRowAttachment() { Name=fileName, Src= $"{destfileFolder}{fileName}", Overwrite=true });
                }
                catch (Exception ex)
                {
                    scope.LogError(ex, "Extract of ListAttachment {0} failed", destfileFolder);
                }
            }
        }

        /// <summary>
        /// Tokenize values of different FieldTypes to string 
        /// </summary>
        /// <param name="web"></param>
        /// <param name="field"></param>
        /// <param name="fieldValue"></param>
        /// <param name="fieldValuesAsText"></param>
        /// <returns></returns>
        private string TokenizeValue(Web web, SPField field, KeyValuePair<string, object> fieldValue, Dictionary<string, string> fieldValuesAsText)
        {
            string value = string.Empty;
            switch (field.TypeAsString)
            {
                case "URL":
                    value = Tokenize(fieldValuesAsText[fieldValue.Key], web.Url, web);
                    break;
                case "User":
                    var userFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldUserValue;
                    if (userFieldValue != null)
                    {
                        value = userFieldValue.Email;
                    }
                    break;
                case "UserMulti":
                    var userMulitFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldUserValue[];
                    if (userMulitFieldValue != null)
                    {
                        value = string.Join(",", userMulitFieldValue.Select(u => u.Email).ToArray())?.TrimEnd(new char[] { ',' });
                    }
                    break;
                case "Lookup":
                    var lookupFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldLookupValue;
                    if (lookupFieldValue != null)
                    {
                        value = lookupFieldValue.LookupId.ToString();
                    }
                    break;
                case "LookupMulti":
                    var lookupMultiFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldLookupValue[];
                    if (lookupMultiFieldValue != null)
                    {
                        value = value = string.Join(",", lookupMultiFieldValue.Select(l => l.LookupId).ToArray())?.TrimEnd(new char[] { ',' });
                    }
                    break;
                case "TaxonomyFieldType":
                    var taxonomyFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue;
                    if (taxonomyFieldValue != null)
                    {
                        value = $"{taxonomyFieldValue.Label}|{taxonomyFieldValue.TermGuid}";
                    }
                    break;
                case "TaxonomyFieldTypeMulti":
                    var taxonomyMultiFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValueCollection;
                    if (taxonomyMultiFieldValue != null)
                    {
                        string terms = "";
                        foreach (var term in taxonomyMultiFieldValue)
                        {
                            terms += $"{term.Label}|{term.TermGuid};";
                        }
                        value = terms.TrimEnd(new char[] { ';' });
                    }
                    break;
                case "DateTime":
                    var dateTimeFieldValue = fieldValue.Value as DateTime?;
                    if (dateTimeFieldValue.HasValue)
                    {
                        value = dateTimeFieldValue.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
                    }
                    break;
                case "ContentTypeIdFieldType":
                default:
                    value = Tokenize(fieldValue.Value.ToString(), web.Url, web);
                    break;
            }

            return value;
        }

        /// <summary>
        /// Extract FieldValues, PropertyBag and Security Settings from Folder
        /// </summary>
        /// <param name="web"></param>
        /// <param name="serverRelativePathToFolder"></param>
        /// <param name="siteSecurity"></param>
        /// <param name="scope"></param>
        /// <param name="fieldsToExclude"></param>
        /// <returns></returns>
        public PnPFolder ExtractFolderSettings(Web web, string serverRelativePathToFolder, PnPSiteSecurity siteSecurity, PnPMonitoredScope scope, string[] fieldsToExclude = null)
        {
            PnPFolder pnpFolder = null;
            try
            {
                SPFolder spFolder = web.GetFolderByServerRelativeUrl(serverRelativePathToFolder);
                web.Context.Load(spFolder,
                    f => f.Name,
                    f => f.ServerRelativeUrl,
                    f => f.Properties,
                    f => f.ListItemAllFields,
                    f => f.ListItemAllFields.RoleAssignments,
                    f => f.ListItemAllFields.RoleAssignments.Include(r => r.Member, r => r.RoleDefinitionBindings),
                    f => f.ListItemAllFields.HasUniqueRoleAssignments,
                    f => f.ListItemAllFields.ParentList,
                    f => f.ListItemAllFields.ContentType.StringId);
                web.Context.Load(web,
                    w => w.AssociatedOwnerGroup,
                    w => w.AssociatedMemberGroup,
                    w => w.AssociatedVisitorGroup,
                    w => w.Title,
                    w => w.Url,
                    w => w.RoleDefinitions.Include(r => r.RoleTypeKind, r => r.Name),
                    w => w.ContentTypes.Include(c => c.Id, c => c.Name, c => c.StringId));
                web.Context.ExecuteQueryRetry();

                pnpFolder = new PnPFolder(spFolder.Name);

                //export PnPFolder Properties
                if (spFolder.Properties.FieldValues.Any())
                {
                    foreach (var propKey in spFolder.Properties.FieldValues.Keys.Where(k => !k.StartsWith("vti_") && !k.StartsWith("docset_")))
                    {
                        pnpFolder.PropertyBagEntries.Add(new PropertyBagEntry() { Key = propKey, Value = spFolder.Properties.FieldValues[propKey].ToString() });
                    }
                }

                //export PnPFolder FieldValues
                if (spFolder.ListItemAllFields.FieldValues.Any())
                {
                    var list = spFolder.ListItemAllFields.ParentList;

                    var fields = list.Fields;
                    web.Context.Load(fields, fs => fs.IncludeWithDefaultProperties(f => f.TypeAsString, f => f.InternalName, f => f.Title));
                    web.Context.ExecuteQueryRetry();

                    var fieldValues = spFolder.ListItemAllFields.FieldValues;

                    var fieldValuesAsText = spFolder.ListItemAllFields.EnsureProperty(li => li.FieldValuesAsText).FieldValues;

                    if (fieldsToExclude == null)
                    {
                        fieldsToExclude = new string[] { };
                    }

                    #region //**** get correct Content Type
                    string ctId = string.Empty;
                    foreach (var ct in web.ContentTypes.OrderByDescending(c => c.StringId.Length))
                    {
                        if (spFolder.ListItemAllFields.ContentType.StringId.StartsWith(ct.StringId))
                        {
                            pnpFolder.ContentTypeID = ct.StringId;
                            break;
                        }
                    }
                    #endregion //**** get correct Content Type

                    foreach (var fieldValue in fieldValues.Where(f => !fieldsToExclude.Contains(f.Key)))
                    {
                        if (fieldValue.Value != null && !string.IsNullOrEmpty(fieldValue.Value.ToString()))
                        {
                            var field = fields.FirstOrDefault(fs => fs.InternalName == fieldValue.Key);
                            string value = string.Empty;

                            //ignore read only fields
                            if (!field.ReadOnlyField || WriteableReadOnlyField.Contains(field.InternalName.ToLower()))
                            {
                                value = TokenizeValue(web, field, fieldValue, fieldValuesAsText);
                            }

                            if (fieldValue.Key.Equals("ContentTypeId",StringComparison.CurrentCultureIgnoreCase))
                            {
                                value = null; //ignore here since already in dataRow
                            }

                            if (fieldValue.Key.Equals("HTML_x0020_File_x0020_Type", StringComparison.CurrentCultureIgnoreCase) &&
                                fieldValuesAsText["HTML_x0020_File_x0020_Type"] == "OneNote.Notebook")
                            {
                                pnpFolder.Properties.Add("File_x0020_Type", "OneNote.Notebook");
                                pnpFolder.Properties.Add(fieldValue.Key, "OneNote.Notebook");
                                value = null;
                            }

                            // We process real values only
                            if (!string.IsNullOrWhiteSpace(value) && value != "[]")
                            {
                                pnpFolder.Properties.Add(fieldValue.Key, value);
                            }
                        }
                    }
                }

                // get PnPFolder Permissions
                if (spFolder.ListItemAllFields.HasUniqueRoleAssignments) // && siteSecurity != null)
                {
                    GetObjectSecurity(web, spFolder.ListItemAllFields.RoleAssignments, pnpFolder.Security);
                }
            }
            catch (Exception ex)
            {
                scope.LogError(ex, "Extract of Folder {0} failed", serverRelativePathToFolder);
            }
            return pnpFolder;
        }

        private void GetObjectSecurity(Web web, Microsoft.SharePoint.Client.RoleAssignmentCollection RoleAssignments, ObjectSecurity objectSecurity)
        {
            objectSecurity.ClearSubscopes = true;
            objectSecurity.CopyRoleAssignments = false;
            var LimitedAccess = web.RoleDefinitions.FirstOrDefault(role => role.RoleTypeKind == RoleType.Guest);

            foreach (var roleAssignment in RoleAssignments)
            {
                foreach (var rDef in roleAssignment.RoleDefinitionBindings.OrderBy(d => d.Order))
                {
                    if (!(LimitedAccess != null && (rDef.Name.Equals(LimitedAccess.Name)))&&!roleAssignment.Member.LoginName.StartsWith("SharingLinks."))
                    {
                        string principalName = roleAssignment.Member.LoginName.Replace(web.Title, "{sitetitle}");
                        //check if we have this Group in Template Security - if so, we add it

                        //if ((!string.IsNullOrWhiteSpace(siteSecurity.AssociatedOwnerGroup) && siteSecurity.AssociatedOwnerGroup.Equals(principalName)) |
                        //    (!string.IsNullOrWhiteSpace(siteSecurity.AssociatedMemberGroup) && siteSecurity.AssociatedMemberGroup.Equals(principalName)) |
                        //    (!string.IsNullOrWhiteSpace(siteSecurity.AssociatedVisitorGroup) && siteSecurity.AssociatedVisitorGroup.Equals(principalName)) |
                        //    siteSecurity.SiteGroups.Any(g => g.Title.Equals(principalName)))
                        //{
                        objectSecurity.RoleAssignments.Add(new PnPRoleAssignment()
                        {
                            Principal = principalName,
                            Remove = false,
                            RoleDefinition = rDef.Name
                        });
                        //}
                    }
                }
            }
        }

        #region //**** static Field Lists as Hint how to handle 
        public static string[] WriteableReadOnlyField = new[]
        {
            "description","publishingpagelayout", "contenttypeid","bannerimageurl","_originalsourceitemid","_originalsourcelistid","_originalsourcesiteid","_originalsourcewebid","_originalsourceurl"
        };

        //ignore this fields on one.note folder
        public static string[] FolderFieldsToExclude = new[] {
                "_Dirty",
                "_IsCurrentVersion",
                "_Level",
                "_ModerationStatus",
                "_Parsable",
                "_UIVersion",
                "_UIVersionString",
                "AppAuthor",
                "AppEditor",
                "Author",
                "ContentTypeId",
                "ContentVersion",
                "Created",
                "Created_x0020_Date",
                "DocConcurrencyNumber",
                "Editor",
                "FileRef",
                "FileLeafRef",
                "FolderChildCount",
                "FSObjType",
                "GUID",
                "ID",
                "IsCheckedoutToLocal",
                "ItemChildCount",
                "Last_x0020_Modified",
                "MetaInfo",
                "Modified",
                "Modified_x0020_By",
                "NoExecute",
                "Order",
                "owshiddenversion",
                "ParentUniqueId",
                "ProgId",
                "ScopeId",
                "SMLastModifiedDate",
                "SMTotalFileStreamSize",
                "SortBehavior",
                "Title",
                "UniqueId",
                "WorkflowVersion",
                "xd_Signature"
            };

        public static string[] FileFieldsToExclude = new[] {
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
                    "StreamHash",
                    "TaxCatchAll",
                    "TaxCatchAllLabel",
                    "_ModerationStatus",
                    //"HtmlDesignAssociated",
                    //"HtmlDesignStatusAndPreview",
                    "MetaInfo",
                    "CheckoutUser",
                    "NoExecute",
                    "_HasCopyDestinations",
                    "ContentVersion",
                    "UIVersion",
                    "Title",
                    "ContentTypeId",
                    //Feld welches sonst doppelt vorhanden wäre
                };

        public static string[] ExcludeCustomListFields = new[]
        {
            "_VirusStatus", "_VirusVendorID", "_VirusInfo", "FileLeafRef", "Attachments"
        };

        public static string[] WriteableReadOnlyListField = new[]
        {
            "description"
        };

        #endregion //**** static Field Lists as Hint how to handle 

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Lists.Any(l => l.DataRows.Any());
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = creationInfo.HandlersToProcess.HasFlag(Handlers.Fields) && creationInfo.HandlersToProcess.HasFlag(Handlers.ContentTypes) && creationInfo.HandlersToProcess.HasFlag(Handlers.Lists);
            }
            return _willExtract.Value;
        }
    }
}