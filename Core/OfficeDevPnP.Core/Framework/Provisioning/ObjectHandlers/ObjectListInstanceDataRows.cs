using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using Field = Microsoft.SharePoint.Client.Field;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectListInstanceDataRows : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "List instances Data Rows"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Lists.Any())
                {
                    var rootWeb = (web.Context as ClientContext).Site.RootWeb;

                    web.EnsureProperties(w => w.ServerRelativeUrl);

                    web.Context.Load(web.Lists, lc => lc.IncludeWithDefaultProperties(l => l.RootFolder.ServerRelativeUrl));
                    web.Context.ExecuteQueryRetry();
                    var existingLists = web.Lists.AsEnumerable<List>().Select(existingList => existingList.RootFolder.ServerRelativeUrl).ToList();
                    var serverRelativeUrl = web.ServerRelativeUrl;

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
                            web.Context.Load(fields, fs => fs.Include(f => f.InternalName, f=>f.Title,f => f.FieldTypeKind, f => f.TypeAsString, f=>f.ReadOnlyField));
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

                                    bool create = true;
                                    ListItem listitem = null;
                                    if (!string.IsNullOrEmpty(listInstance.DataRows.KeyColumn))
                                    {
                                        // Get value from key column
                                        var dataRowValues = dataRow.Values.Where(v => v.Key == listInstance.DataRows.KeyColumn);

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
                                                    create = false;
                                                }
                                                else
                                                {
                                                    listitem = existingItems[0];
                                                    create = true;
                                                }
                                            }
                                        }
                                    }

                                    if (create)
                                    {
                                        if (listitem == null)
                                        {
                                            var listitemCI = new ListItemCreationInformation();
                                            listitem = list.AddItem(listitemCI);
                                        }

                                        foreach (var dataValue in dataRow.Values)
                                        {
                                            string fieldName = parser.ParseString(dataValue.Key);
                                            Field dataField = fields.FirstOrDefault(
                                                f => f.InternalName == fieldName);

                                            if (dataField == null)
                                            {
                                                scope.LogWarning("Cannot find field {0} in list {1}", fieldName, listInstance.Title);
                                            }
                                            else if (dataField.ReadOnlyField)
                                            {
                                                scope.LogWarning("Cannot set field {0} in list {1} because it's a readonly field", fieldName, listInstance.Title);
                                            }
                                            else
                                            {
                                                String fieldValue = parser.ParseString(dataValue.Value);
                                                scope.LogDebug("Setting field {0} in list {1} with value {2}", dataField.Title, listInstance.Title, fieldValue);
                                                try
                                                {
                                                    switch (dataField.FieldTypeKind)
                                                    {
                                                        case FieldType.Geolocation:
                                                            if (fieldValue == null)
                                                            {
                                                                listitem[fieldName] = null;
                                                                break;
                                                            }
                                                            // FieldGeolocationValue - Expected format: Altitude,Latitude,Longitude,Measure
                                                            var geolocationArray = fieldValue.Split(',');
                                                            if (geolocationArray.Length == 4)
                                                            {
                                                                var geolocationValue = new FieldGeolocationValue
                                                                {
                                                                    Altitude = Double.Parse(geolocationArray[0]),
                                                                    Latitude = Double.Parse(geolocationArray[1]),
                                                                    Longitude = Double.Parse(geolocationArray[2]),
                                                                    Measure = Double.Parse(geolocationArray[3]),
                                                                };
                                                                listitem[fieldName] = geolocationValue;
                                                            }
                                                            else
                                                            {
                                                                listitem[fieldName] = fieldValue;
                                                            }
                                                            break;

                                                        case FieldType.Lookup:
                                                            if (fieldValue == null)
                                                            {
                                                                listitem[fieldName] = null;
                                                                break;
                                                            }
                                                            // FieldLookupValue - Expected format: LookupID or LookupID,LookupID,LookupID...
                                                            if (fieldValue.Contains(","))
                                                            {
                                                                var lookupValues = new List<FieldLookupValue>();
                                                                fieldValue.Split(',').All(value =>
                                                                {
                                                                    lookupValues.Add(new FieldLookupValue
                                                                    {
                                                                        LookupId = int.Parse(value),
                                                                    });
                                                                    return true;
                                                                });
                                                                listitem[fieldName] = lookupValues.ToArray();
                                                            }
                                                            else
                                                            {
                                                                var lookupValue = new FieldLookupValue
                                                                {
                                                                    LookupId = int.Parse(fieldValue),
                                                                };
                                                                listitem[fieldName] = lookupValue;
                                                            }
                                                            break;

                                                        case FieldType.URL:
                                                            if (fieldValue == null)
                                                            {
                                                                listitem[fieldName] = null;
                                                                break;
                                                            }

                                                            // FieldUrlValue - Expected format: URL,Description
                                                            var urlArray = fieldValue.Split(',');
                                                            var linkValue = new FieldUrlValue();
                                                            if (urlArray.Length == 2)
                                                            {
                                                                linkValue.Url = urlArray[0];
                                                                linkValue.Description = urlArray[1];
                                                            }
                                                            else
                                                            {
                                                                linkValue.Url = urlArray[0];
                                                                linkValue.Description = urlArray[0];
                                                            }
                                                            listitem[fieldName] = linkValue;
                                                            break;

                                                        case FieldType.User:
                                                            if (fieldValue == null)
                                                            {
                                                                listitem[fieldName] = null;
                                                                break;
                                                            }

                                                            // FieldUserValue - Expected format: loginName or loginName,loginName,loginName...
                                                            if (fieldValue.Contains(","))
                                                            {
                                                                var userValues = new List<FieldUserValue>();
                                                                fieldValue.Split(',').All(value =>
                                                                {
                                                                    var user = web.EnsureUser(value);
                                                                    web.Context.Load(user);
                                                                    web.Context.ExecuteQueryRetry();
                                                                    if (user != null)
                                                                    {
                                                                        userValues.Add(new FieldUserValue
                                                                        {
                                                                            LookupId = user.Id,
                                                                        }); ;
                                                                    }
                                                                    return true;
                                                                });
                                                                listitem[fieldName] = userValues.ToArray();
                                                            }
                                                            else
                                                            {
                                                                var user = web.EnsureUser(fieldValue);
                                                                web.Context.Load(user);
                                                                web.Context.ExecuteQueryRetry();
                                                                if (user != null)
                                                                {
                                                                    var userValue = new FieldUserValue
                                                                    {
                                                                        LookupId = user.Id,
                                                                    };
                                                                    listitem[fieldName] = userValue;
                                                                }
                                                                else
                                                                {
                                                                    listitem[fieldName] = fieldValue;
                                                                }
                                                            }
                                                            break;

                                                        case FieldType.DateTime:
                                                            if (DateTime.TryParse(fieldValue, out DateTime dateTime))
                                                            {
                                                                listitem[fieldName] = dateTime;
                                                            }
                                                            else
                                                            {
                                                                listitem[fieldName] = null;
                                                            }
                                                            break;
                                                        case FieldType.Attachments:
                                                        case FieldType.Computed:
                                                            scope.LogWarning("Unsupported field type: {0} (Field:{1})", dataField.FieldTypeKind, fieldName);
                                                            break;
                                                        default:
                                                            listitem[fieldName] = fieldValue;
                                                            break;
                                                    }
                                                    listitem.Update();
                                                }
                                                catch (Exception exc)
                                                {
                                                    scope.LogWarning("Error when reading value {0} for field {1}. Exception : {2}", fieldValue, fieldName, exc.Message);
                                                    scope.LogDebug("Exception was {0}", exc);
                                                }
                                            }
                                        }
                                        web.Context.ExecuteQueryRetry(); // TODO: Run in batches?

                                        if (dataRow.Security != null && (dataRow.Security.ClearSubscopes == true || dataRow.Security.CopyRoleAssignments == true || dataRow.Security.RoleAssignments.Count > 0))
                                        {
                                            listitem.SetSecurity(parser, dataRow.Security);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (ex.GetType().Equals(typeof(ServerException)) &&
                                        (ex as ServerException).ServerErrorTypeName.Equals("Microsoft.SharePoint.SPDuplicateValuesFoundException", StringComparison.InvariantCultureIgnoreCase) &&
                                        applyingInformation.IgnoreDuplicateDataRowErrors)
                                    {
                                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_duplicate);
                                        continue;
                                    }
                                    else
                                    {
                                        scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_failed___0_____1_, ex.Message, ex.StackTrace);
                                        throw;
                                    }
                                }
                            }
                        }
                    }

                    #endregion DataRows
                }
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