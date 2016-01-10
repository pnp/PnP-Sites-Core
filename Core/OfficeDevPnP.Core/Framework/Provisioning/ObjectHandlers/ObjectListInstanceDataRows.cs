using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Field = Microsoft.SharePoint.Client.Field;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;

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
                            var list = web.Lists.GetByTitle(listInstance.Title);
                            web.Context.Load(list);

                            // Retrieve the fields' types from the list
                            Microsoft.SharePoint.Client.FieldCollection fields = list.Fields;
                            web.Context.Load(fields, fs => fs.Include(f => f.InternalName, f => f.FieldTypeKind));
                            web.Context.ExecuteQueryRetry();

                            foreach (var dataRow in listInstance.DataRows)
                            {
                                try
                                {
                                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_list_item__0_, listInstance.DataRows.IndexOf(dataRow) + 1);
                                    var listitemCI = new ListItemCreationInformation();
                                    var listitem = list.AddItem(listitemCI);

                                    foreach (var dataValue in dataRow.Values)
                                    {
                                        Field dataField = fields.FirstOrDefault(
                                            f => f.InternalName == parser.ParseString(dataValue.Key));

                                        if (dataField != null)
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
                                                        listitem[parser.ParseString(dataValue.Key)] = geolocationValue;
                                                    }
                                                    else
                                                    {
                                                        listitem[parser.ParseString(dataValue.Key)] = fieldValue;
                                                    }
                                                    break;
                                                case FieldType.Lookup:
                                                    // FieldLookupValue - Expected format: LookupID
                                                    var lookupValue = new FieldLookupValue
                                                    {
                                                        LookupId = Int32.Parse(fieldValue),
                                                    };
                                                    listitem[parser.ParseString(dataValue.Key)] = lookupValue;
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
                                                    listitem[parser.ParseString(dataValue.Key)] = linkValue;
                                                    break;
                                                case FieldType.User:
                                                    // FieldUserValue - Expected format: loginName
                                                    var user = web.EnsureUser(fieldValue);
                                                    web.Context.Load(user);
                                                    web.Context.ExecuteQueryRetry();

                                                    if (user != null)
                                                    {
                                                        var userValue = new FieldUserValue
                                                        {
                                                            LookupId = user.Id,
                                                        };
                                                        listitem[parser.ParseString(dataValue.Key)] = userValue;
                                                    }
                                                    else
                                                    {
                                                        listitem[parser.ParseString(dataValue.Key)] = fieldValue;
                                                    }
                                                    break;
                                                default:
                                                    listitem[parser.ParseString(dataValue.Key)] = fieldValue;
                                                    break;
                                            }
                                        }
                                        listitem.Update();
                                    }
                                    web.Context.ExecuteQueryRetry(); // TODO: Run in batches?
                                    
                                    if (dataRow.Security != null && dataRow.Security.RoleAssignments.Count != 0)
                                    {
                                        listitem.SetSecurity(parser, dataRow.Security);
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

        public override bool WillProvision(Web web, ProvisioningTemplate template)
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

