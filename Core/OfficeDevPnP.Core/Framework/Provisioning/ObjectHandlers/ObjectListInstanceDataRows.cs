using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
                                    bool IsNewItem = false;
                                    if (listitem == null)
                                    {
                                        var listitemCI = new ListItemCreationInformation();
                                        listitem = list.AddItem(listitemCI);
                                        IsNewItem = true;
                                    }

                                    ListItemUtilities.UpdateListItem(listitem, parser, dataRow.Values, ListItemUtilities.ListItemUpdateType.UpdateOverwriteVersion, IsNewItem);

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
                                            if (!IsNewItem)
                                            {
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
                                            else
                                            {
                                                AddAttachment(template, listitem, attachment, IsNewItem);
                                            }
                                        }
                                        if (IsNewItem)
                                            listitem.Context.ExecuteQueryRetry();
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


        private void AddAttachment(ProvisioningTemplate template, ListItem listitem, Model.SharePoint.InformationArchitecture.DataRowAttachment attachment, bool SkipExecuteQuery = false)
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
            if (!SkipExecuteQuery)
                listitem.Context.ExecuteQueryRetry();
            else
                listitem.Update();
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var lists = web.Lists;
                web.EnsureProperty(w => w.ServerRelativeUrl);
                web.Context.Load(lists,
                  lc => lc.IncludeWithDefaultProperties(
                        l => l.RootFolder.ServerRelativeUrl,
                        l => l.Fields.IncludeWithDefaultProperties(
                          f => f.Id,
                          f => f.Title,
                          f => f.Hidden,
                          f => f.InternalName,
                          f => f.DefaultValue,
                          f => f.Required))
                  );
                web.Context.ExecuteQueryRetry();

                var allLists = new List<List>();

                var listsToProcess = lists.AsEnumerable().Where(l => l.Hidden == false || l.Hidden == creationInfo.IncludeHiddenLists).ToArray();
                var listCount = 0;
                foreach (var siteList in listsToProcess)
                {
                    if (!creationInfo.ListsExtractionConfiguration.Any(i =>
                      {
                          Guid listId;
                          if (Guid.TryParse(i.Title, out listId))
                          {
                              return (listId == siteList.Id);
                          }
                          else
                          {
                              return (false);
                          }
                      }) && creationInfo.ListsExtractionConfiguration.FirstOrDefault(i => i.Title.Equals(siteList.Title) && i.IncludeItems) == null)
                    {
                        continue;
                    }
                    var extractionInfo = creationInfo.ListsExtractionConfiguration.FirstOrDefault(e => e.Title.Equals(siteList.Title));
                    CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();
                    ExtractConfiguration.ExtractListsQueryConfiguration queryConfig = null;
                    if (extractionInfo.Query != null)
                    {
                        queryConfig = extractionInfo.Query;

                        camlQuery = new CamlQuery();

                        string viewXml = $"<View><Query>{queryConfig.CamlQuery}</Query>";
                        if (queryConfig.ViewFields != null && queryConfig.ViewFields.Count > 0)
                        {
                            viewXml += "<ViewFields>";
                            foreach (var viewField in queryConfig.ViewFields)
                            {
                                viewXml += $"<FieldRef Name='{viewField}' />";
                            }
                            viewXml += "</ViewFields>";
                        }
                        if (queryConfig.RowLimit > 0)
                        {
                            viewXml += $"<RowLimit>{queryConfig.RowLimit}</RowLimit>";
                        }
                        viewXml += "</View>";
                        camlQuery.ViewXml = viewXml;

                    }
                    var listInstance = template.Lists.FirstOrDefault(l => siteList.RootFolder.ServerRelativeUrl.Equals(UrlUtility.Combine(web.ServerRelativeUrl, l.Url)));
                    if (listInstance != null)
                    {

                        var items = siteList.GetItems(camlQuery);
                        siteList.Context.Load(items, i => i.Include(li => li.FieldValuesAsText));
                        if(queryConfig != null && queryConfig.ViewFields.Any())
                        {
                            foreach (var viewField in queryConfig.ViewFields)
                            {
                                siteList.Context.Load(items, i => i.Include(li => li[viewField]));
                            }
                        }
                        siteList.Context.ExecuteQueryRetry();
                        foreach (var item in items)
                        {
                            var dataRow = new Model.DataRow();
                            foreach (var fieldValue in item.FieldValues)
                            {
                                // var isInternal = false;
                                //var listField = siteList.Fields.FirstOrDefault(f => f.InternalName.Equals(fieldValue.Key));
                                //if (listField != null)
                                //{
                                //    isInternal = BuiltInFieldId.Contains(listField.Id);
                                //}
                                var value = item.FieldValuesAsText[fieldValue.Key];
                                var skip = extractionInfo.SkipEmptyFields && string.IsNullOrEmpty(value);
                                if (!skip)
                                {
                                    dataRow.Values.Add(fieldValue.Key, value);
                                }
                            }
                            listInstance.DataRows.Add(dataRow);
                        }
                    }
                }
            }
            return template;
        }

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
                _willExtract = creationInfo.ListsExtractionConfiguration != null && creationInfo.ListsExtractionConfiguration.Count > 0;
            }
            return _willExtract.Value;
        }
    }
}