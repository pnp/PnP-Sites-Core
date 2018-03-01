using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using SPField = Microsoft.SharePoint.Client.Field;
using SPFieldCollection = Microsoft.SharePoint.Client.FieldCollection;
using SPTerm = Microsoft.SharePoint.Client.Taxonomy.Term;
namespace OfficeDevPnP.Core.Extensions
{
    public static class ListItemExtensions
    {
        public static void SetListItemValue(
            this ListItem listitem,
            Web web,
            SPField dataField,
            string dataValue,
            TokenParser parser,
            PnPMonitoredScope scope
            )
        {
            if (listitem == null) throw new ArgumentNullException(nameof(listitem));
            if (web == null) throw new ArgumentNullException(nameof(web));
            if (dataField == null) throw new ArgumentNullException(nameof(dataField));

            var fieldName = dataField.InternalName;
            var fieldValue = parser.ParseString(dataValue);
            switch (dataField.FieldTypeKind)
            {
                case FieldType.Geolocation:
                    {
                        if (fieldValue == null)
                        {
                            listitem[fieldName] = null;
                            break;
                        }
                        // FieldGeolocationValue - Expected format: Altitude,Latitude,Longitude,Measure
                        var geolocationArray = fieldValue.Split(',');
                        if (geolocationArray.Length == 4)
                        {
                            listitem[fieldName] = new FieldGeolocationValue
                            {
                                Altitude = Double.Parse(geolocationArray[0]),
                                Latitude = Double.Parse(geolocationArray[1]),
                                Longitude = Double.Parse(geolocationArray[2]),
                                Measure = Double.Parse(geolocationArray[3]),
                            };
                        }
                        else
                        {
                            listitem[fieldName] = fieldValue;
                        }
                        break;
                    }

                case FieldType.Lookup:
                    {
                        if (fieldValue == null)
                        {
                            listitem[fieldName] = null;
                            break;
                        }
                            // FieldLookupValue - Expected format: LookupID or LookupID,LookupID,LookupID...
                            var lookupField = web.Context.CastTo<FieldLookup>(dataField);
                            if (lookupField.AllowMultipleValues)
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
                                listitem[fieldName] = new FieldLookupValue
                                {
                                    LookupId = int.Parse(fieldValue),
                                };
                            }

                        break;
                    }

                case FieldType.URL:
                    {
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
                    }

                case FieldType.User:
                    {
                        if (fieldValue == null)
                        {
                            listitem[fieldName] = null;
                            break;
                        }

                        // FieldUserValue - Expected format: loginName or loginName,loginName,loginName...
                        var userField = web.Context.CastTo<FieldUser>(dataField);
                        if (userField.AllowMultipleValues)
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
                                    });
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
                                listitem[fieldName] = new FieldUserValue
                                {
                                    LookupId = user.Id,
                                };
                            }
                            else
                            {
                                listitem[fieldName] = fieldValue;
                            }
                        }
                        break;
                    }

                case FieldType.DateTime:
                    {
                        if (DateTime.TryParse(fieldValue, out DateTime dateTime))
                        {
                            listitem[fieldName] = dateTime;
                        }
                        else
                        {
                            listitem[fieldName] = null;
                        }
                        break;
                    }

                case FieldType.MultiChoice:
                    if (string.IsNullOrWhiteSpace(fieldValue))
                    {
                        listitem[fieldName] = null;
                        break;
                    }

                    if (JsonUtility.TryDeserialize(fieldValue, out string[] choiceValues))
                    {
                        // Backward compatibility. See https://github.com/SharePoint/PnP-Sites-Core/issues/1577
                        listitem[fieldName] = choiceValues;
                    }
                    else
                    {
                        listitem[fieldName] = fieldValue;
                    }
                    break;
                case FieldType.Invalid:
                    if (dataField.TypeAsString == "TaxonomyFieldType" || dataField.TypeAsString == "TaxonomyFieldTypeMulti")
                    {
                        var taxonomyField = web.Context.CastTo<TaxonomyField>(dataField);
                        taxonomyField.EnsureProperties(tf => tf.AllowMultipleValues, tf=>tf.IsKeyword, tf=>tf.TermSetId);
                        var site = ((ClientContext)web.Context).Site;
                        if (taxonomyField.AllowMultipleValues)
                        {
                            if (JsonUtility.TryDeserialize(fieldValue, out TaxonomyFieldValue[] taxValue))
                            {
                                var termValues = taxValue.ToDictionary(tv => new Guid(tv.TermGuid), tv => tv.Label);
                                listitem.SetTaxonomyFieldValues(taxonomyField.Id, termValues);
                            }
                            else
                            {
                                var terms = new Dictionary<Guid, String>();
                                foreach (var termLabel in fieldValue.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                                {
                                    var term = site.GetTermByName(taxonomyField.TermSetId, termLabel);
                                    if(term != null)
                                    {
                                        terms.Add(term.Id, termLabel);
                                    }
                                    else
                                    {
                                        scope?.LogWarning("Term {0} is not found (setting field {1})", termLabel, dataField.InternalName);
                                    }
                                }

                                listitem.SetTaxonomyFieldValues(taxonomyField.Id, terms);
                            }
                        }
                        else
                        {
                            if (JsonUtility.TryDeserialize(fieldValue, out TaxonomyFieldValue taxValue))
                            {
                                taxonomyField.SetFieldValueByValue(listitem, taxValue);
                            }
                            else
                            {
                                var term = site.GetTermByName(taxonomyField.TermSetId, fieldValue);
                                if (term != null)
                                {
                                    taxonomyField.SetFieldValueByTerm(listitem, term, (int)web.Language);
                                }
                                else
                                {
                                    scope?.LogWarning("Term {0} is not found (setting field {1})", fieldValue, dataField.InternalName);
                                }
                            }
                        }
                    }
                    else
                    {
                        scope?.LogWarning("Unsupported field type: {0}/ (Field:{1})", dataField.FieldTypeKind, dataField.TypeAsString, fieldName);

                    }
                    break;
                case FieldType.Attachments:
                case FieldType.Computed:
                    {
                        scope.LogWarning("Unsupported field type: {0} (Field:{1})", dataField.FieldTypeKind, fieldName);
                        break;
                    }
                default:
                    {
                        listitem[fieldName] = fieldValue;
                        break;
                    }
            }
        }


        public static void SetListItemFieldValue(
            this ListItem listitem,
            Web web,
            TokenParser parser,
            PnPMonitoredScope scope,
            ListInstance listInstance,
            SPFieldCollection fields,
            string fieldName,
            string valueAsAsString
            )
        {
            SPField dataField = fields.FirstOrDefault(
                f => f.InternalName == fieldName
                );

            if (dataField == null)
            {
                scope?.LogWarning("Cannot find field {0} in list {1}", fieldName, listInstance.Title);
            }
            else if (dataField.ReadOnlyField)
            {
                scope?.LogWarning("Cannot set field {0} in list {1} because it's a readonly field", fieldName, listInstance.Title);
            }
            else
            {
                var fieldValue = parser.ParseString(valueAsAsString);
                scope?.LogDebug("Setting field {0} in list {1} with value {2}", dataField.Title, listInstance.Title, fieldValue);
                try
                {
                    listitem.SetListItemValue(web, dataField, fieldValue, parser, scope);
                }
                catch (Exception exc)
                {
                    scope?.LogWarning("Error when reading value {0} for field {1}. Exception : {2}", fieldValue, fieldName, exc.Message);
                    scope?.LogDebug("Exception was {0}", exc);
                }
            }
        }
    }
}