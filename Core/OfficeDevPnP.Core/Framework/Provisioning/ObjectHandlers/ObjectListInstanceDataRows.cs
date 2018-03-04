using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Field = Microsoft.SharePoint.Client.Field;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using Microsoft.SharePoint.Client.Taxonomy;

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
                                    var updateValues = new List<FieldUpdateValue>();

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

                                        foreach (var dataValue in dataRow.Values)
                                        {

                                            Field dataField = fields.FirstOrDefault(
                                                f => f.InternalName == parser.ParseString(dataValue.Key));

                                            if (dataField != null && dataField.ReadOnlyField)
                                            {
                                                // skip read only fields
                                                continue;
                                            }
                                            if (dataField != null)
                                            {
                                                if (dataValue.Value == null)
                                                {
                                                    if (dataField.FieldTypeKind == FieldType.Invalid)
                                                    {
                                                        updateValues.Add(new FieldUpdateValue(dataValue.Key, null, dataField.TypeAsString));
                                                    }
                                                    else
                                                    {
                                                        updateValues.Add(new FieldUpdateValue(dataValue.Key, null));
                                                    }
                                                }
                                                else
                                                {
                                                    String fieldValue = parser.ParseString(dataValue.Value);

                                                    switch (dataField.FieldTypeKind)
                                                    {
                                                        case FieldType.Geolocation:
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
                                                                updateValues.Add(new FieldUpdateValue(dataValue.Key, geolocationValue));
                                                            }
                                                            else
                                                            {
                                                                updateValues.Add(new FieldUpdateValue(dataValue.Key, fieldValue));
                                                            }
                                                            break;
                                                        case FieldType.Lookup:
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
                                                                updateValues.Add(new FieldUpdateValue(dataValue.Key, lookupValues.ToArray()));
                                                            }
                                                            else
                                                            {
                                                                var lookupValue = new FieldLookupValue
                                                                {
                                                                    LookupId = int.Parse(fieldValue),
                                                                };
                                                                updateValues.Add(new FieldUpdateValue(dataValue.Key, lookupValue));
                                                            }
                                                            break;
                                                        case FieldType.URL:
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
                                                            updateValues.Add(new FieldUpdateValue(dataValue.Key, linkValue));
                                                            break;
                                                        case FieldType.User:
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
                                                                updateValues.Add(new FieldUpdateValue(dataValue.Key, userValues.ToArray()));
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
                                                                    updateValues.Add(new FieldUpdateValue(dataValue.Key, userValue));
                                                                }
                                                                else
                                                                {
                                                                    updateValues.Add(new FieldUpdateValue(dataValue.Key, fieldValue));
                                                                }
                                                            }
                                                            break;
                                                        case FieldType.DateTime:
                                                            var dateTime = DateTime.MinValue;
                                                            if (DateTime.TryParse(fieldValue, out dateTime))
                                                            {
                                                                updateValues.Add(new FieldUpdateValue(dataValue.Key, dateTime));
                                                            }
                                                            break;
                                                        case FieldType.Invalid:
                                                            switch (dataField.TypeAsString)
                                                            {
                                                                case "TaxonomyFieldType":
                                                                // Single value field - Expected format: term label|term GUID
                                                                case "TaxonomyFieldTypeMulti":
                                                                    // Multi value field - Expected format: term label|term GUID;term label|term GUID;term label|term GUID;...
                                                                    {
                                                                        if (fieldValue != null)
                                                                        {
                                                                            var termStrings = new List<string>();

                                                                            var termsArray = fieldValue.Split(new char[] { ';' });
                                                                            foreach (var term in termsArray)
                                                                            {
                                                                                termStrings.Add($"-1;#{term}");
                                                                            }
                                                                            updateValues.Add(new FieldUpdateValue(dataValue.Key, termStrings, dataField.TypeAsString));
                                                                        }
                                                                        break;
                                                                    }
                                                            }
                                                            break;

                                                        default:
                                                            updateValues.Add(new FieldUpdateValue(dataValue.Key, fieldValue));
                                                            break;
                                                    }
                                                }
                                            }
                                        }

                                        foreach (var itemValue in updateValues.Where(u => u.FieldTypeString != "TaxonomyFieldTypeMulti" && u.FieldTypeString != "TaxonomyFieldType"))
                                        {
                                            if (string.IsNullOrEmpty(itemValue.FieldTypeString))
                                            {
                                                listitem[itemValue.Key] = itemValue.Value;
                                            }
                                        }
                                        listitem.Update();
                                        web.Context.Load(listitem);
                                        web.Context.ExecuteQueryRetry();
                                        var itemId = listitem.Id;
                                        foreach (var itemValue in updateValues.Where(u => u.FieldTypeString == "TaxonomyFieldTypeMulti" || u.FieldTypeString == "TaxonomyFieldType"))
                                        {
                                            switch (itemValue.FieldTypeString)
                                            {
                                                case "TaxonomyFieldTypeMulti":
                                                    {

                                                        var field = fields.FirstOrDefault(f => f.InternalName == itemValue.Key as string || f.Title == itemValue.Key as string);
                                                        var taxField = web.Context.CastTo<TaxonomyField>(field);
                                                        if (itemValue.Value != null)
                                                        {
                                                            var valueCollection = new TaxonomyFieldValueCollection(web.Context, string.Join(";#", itemValue.Value as List<string>), taxField);
                                                            taxField.SetFieldValueByValueCollection(listitem, valueCollection);

                                                        }
                                                        else
                                                        {
                                                            var valueCollection = new TaxonomyFieldValueCollection(web.Context, null, taxField);
                                                            taxField.SetFieldValueByValueCollection(listitem, valueCollection);
                                                        }
                                                        listitem.Update();
                                                        web.Context.Load(listitem);
                                                        web.Context.ExecuteQueryRetry();
                                                        break;
                                                    }
                                                case "TaxonomyFieldType":
                                                    {
                                                        var field = fields.FirstOrDefault(f => f.InternalName == itemValue.Key as string || f.Title == itemValue.Key as string);
                                                        var taxField = web.Context.CastTo<TaxonomyField>(field);
                                                        taxField.EnsureProperty(f => f.TextField);
                                                        var taxValue = new TaxonomyFieldValue();
                                                        if (itemValue.Value != null)
                                                        {
                                                            var termString = (itemValue.Value as List<string>).First();
                                                            taxValue.Label = termString.Split(new string[] { ";#" }, StringSplitOptions.None)[1].Split(new char[] { '|' })[0];
                                                            taxValue.TermGuid = termString.Split(new string[] { ";#" }, StringSplitOptions.None)[1].Split(new char[] { '|' })[1];
                                                            taxValue.WssId = -1;
                                                            taxField.SetFieldValueByValue(listitem, taxValue);
                                                        }
                                                        else
                                                        {
                                                            taxValue.Label = string.Empty;
                                                            taxValue.TermGuid = "11111111-1111-1111-1111-111111111111";
                                                            taxValue.WssId = -1;
                                                            Field hiddenField = list.Fields.GetById(taxField.TextField);
                                                            listitem.Context.Load(hiddenField, tf => tf.InternalName);
                                                            listitem.Context.ExecuteQueryRetry();
                                                            taxField.SetFieldValueByValue(listitem, taxValue); // this order of updates is important.
                                                            listitem[hiddenField.InternalName] = string.Empty; // this order of updates is important.
                                                        }
                                                        listitem.Update();
                                                        web.Context.Load(listitem);
                                                        web.Context.ExecuteQueryRetry();
                                                        break;
                                                    }
                                            }
                                        }
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

                    #endregion
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

        private class FieldUpdateValue
        {
            public string Key { get; set; }
            public object Value { get; set; }
            public string FieldTypeString { get; set; }

            public FieldUpdateValue(string key, object value)
            {
                Key = key;
                Value = value;
            }
            public FieldUpdateValue(string key, object value, string fieldTypeString)
            {
                Key = key;
                Value = value;
                FieldTypeString = fieldTypeString;
            }
        }
    }


}

