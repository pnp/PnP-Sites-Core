using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
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

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            //using (var scope = new PnPMonitoredScope(this.Name))
            //{ }
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
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }
}