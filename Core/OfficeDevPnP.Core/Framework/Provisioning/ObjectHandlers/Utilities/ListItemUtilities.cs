using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    public static class ListItemUtilities
    {
        private readonly static string[] WriteableReadOnlyFields = new string[] { "publishingpagelayout", "contenttypeid" };

        public static FieldUpdateValue ParseFieldValue(Web web, string fieldValue, Field dataField)
        {
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
                        return new FieldUpdateValue(dataField.InternalName, geolocationValue);
                    }
                    else
                    {
                        return new FieldUpdateValue(dataField.InternalName, fieldValue);
                    }
                case FieldType.Lookup:
                    if (dataField.TypeAsString == "LookupMulti" && TryDeserializeAsJson(fieldValue, out FieldLookupValue[] lookupValues2))
                    {
                        // Backward compatibility, when format was stored as json
                        return new FieldUpdateValue(dataField.InternalName, lookupValues2);
                    }
                    // FieldLookupValue - Expected format: LookupID or LookupID,LookupID,LookupID...
                    else if (fieldValue.Contains(","))
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
                        return new FieldUpdateValue(dataField.InternalName, lookupValues.ToArray());
                    }
                    else
                    {
                        var lookupValue = new FieldLookupValue
                        {
                            LookupId = int.Parse(fieldValue),
                        };
                        return new FieldUpdateValue(dataField.InternalName, lookupValue);
                    }
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
                    return new FieldUpdateValue(dataField.InternalName, linkValue);

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
                                });
                            }
                            return true;
                        });
                        return new FieldUpdateValue(dataField.InternalName, userValues.ToArray());
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
                            return new FieldUpdateValue(dataField.InternalName, userValue);
                        }
                        else
                        {
                            return new FieldUpdateValue(dataField.InternalName, fieldValue);
                        }
                    }
                case FieldType.DateTime:
                    var dateTime = DateTime.MinValue;
                    if (DateTime.TryParse(fieldValue, out dateTime))
                    {
                        return new FieldUpdateValue(dataField.InternalName, dateTime);
                    }
                    break;

                case FieldType.MultiChoice:
                    if (TryDeserializeAsJson(fieldValue, out string[] choices))
                    {
                        // Backward compatibility: format is json
                        return new FieldUpdateValue(dataField.InternalName, choices);
                    }
                    else
                    {
                        // expected format: Choice1;#Choice2;#Choice3,
                        return new FieldUpdateValue(dataField.InternalName, fieldValue.Split(";#"));
                    }
                case FieldType.Invalid:
                    switch (dataField.TypeAsString)
                    {
                        case "TaxonomyFieldType":
                            {
                                if (fieldValue != null && TryDeserializeAsJson(fieldValue, out TaxonomyFieldValue taxVal))
                                {
                                    // Backward compatibility, when format was stored as json
                                    return new FieldUpdateValue(
                                        dataField.InternalName,
                                        new List<string> { $"-1;#{taxVal.Label}|{taxVal.TermGuid}" },
                                        dataField.TypeAsString
                                        );
                                }
                                else if (fieldValue != null)
                                {
                                    return new FieldUpdateValue(
                                        dataField.InternalName,
                                        new List<string> { $"-1;#{fieldValue}" },
                                        dataField.TypeAsString
                                        );
                                }
                                break;
                            }
                        // Single value field - Expected format: term label|term GUID
                        case "TaxonomyFieldTypeMulti":
                            {
                                if (fieldValue != null && TryDeserializeAsJson(fieldValue, out TaxonomyFieldValue[] taxValues))
                                {
                                    // Backward compatibility, when format was stored as json
                                    return new FieldUpdateValue(
                                        dataField.InternalName,
                                        taxValues.Select(taxVal=>$"-1;#{taxVal.Label}|{taxVal.TermGuid}"),
                                        dataField.TypeAsString
                                        );

                                }
                                else if (fieldValue != null)
                                {
                                    // Multi value field - Expected format: term label|term GUID;term label|term GUID;term label|term GUID;...
                                    var termStrings = new List<string>();

                                    var termsArray = fieldValue.Split(new char[] { ';' });
                                    foreach (var term in termsArray)
                                    {
                                        termStrings.Add($"-1;#{term}");
                                    }
                                    return new FieldUpdateValue(dataField.InternalName, termStrings, dataField.TypeAsString);
                                }
                                break;
                            }
                    }
                    break;
            }

            // Default to set direct value
            return new FieldUpdateValue(dataField.InternalName, fieldValue, dataField.TypeAsString);
        }

        public static void UpdateListItem(
            Web web,
            ListItem listitem,
            TokenParser parser,
            FieldCollection fields,
            IDictionary<string, string> fieldValues
            )
        {
            var updateValues = new List<FieldUpdateValue>();

            foreach (var dataValue in fieldValues)
            {
                Field dataField = null;

                if (parser != null)
                {
                    dataField = fields.FirstOrDefault(f => f.InternalName == parser.ParseString(dataValue.Key));
                }
                else
                {
                    dataField = fields.FirstOrDefault(f => f.InternalName == dataValue.Key);
                }

                if (dataField == null)
                {
                    // TODO: log Warning
                    continue;
                }

                // Changed by PaoloPia because there are fields like PublishingPageLayout
                // which are marked as read-only, but have to be overwritten while uploading
                // a publishing page file and which in reality can still be written
                if (
                    dataField.ReadOnlyField
                    && !WriteableReadOnlyFields.Contains(dataField.InternalName.ToLower()))
                {
                    // skip read only fields
                    continue;
                }

                if (dataValue.Value == null)
                {
                    updateValues.Add(new FieldUpdateValue(dataValue.Key, null, dataField.TypeAsString));
                }
                else
                {
                    var fieldValue = parser.ParseString(dataValue.Value);

                    updateValues.Add(
                        ParseFieldValue(web, fieldValue, dataField)
                        );
                }
            }

            UpdateListItem(web, listitem, fields, updateValues);
        }

        public static void UpdateListItem(
            Web web,
            ListItem listItem,
            FieldCollection listFields,
            IEnumerable<FieldUpdateValue> updateValues
            )
        {
            if (web == null) throw new ArgumentNullException(nameof(web));
            if (listFields == null) throw new ArgumentNullException(nameof(listFields));
            if (listItem == null) throw new ArgumentNullException(nameof(listItem));

            if (updateValues == null || !updateValues.Any()) return;

            foreach (var itemValue in updateValues.Where(u => u.FieldTypeString != "TaxonomyFieldTypeMulti" && u.FieldTypeString != "TaxonomyFieldType"))
            {
                // Special case for ContentType field
                if (itemValue.Key == "ContentType")
                {
                    var targetCT = listItem.ParentList.GetContentTypeByName((string)itemValue.Value);
                    web.Context.ExecuteQueryRetry();

                    if (targetCT != null)
                    {
                        listItem["ContentTypeId"] = targetCT.StringId;
                    }
                    else
                    {
                        Log.Error(Constants.LOGGING_SOURCE, "Content Type {0} does not exist in target list!", (string)itemValue.Value);
                    }
                }
                else
                {
                    listItem[itemValue.Key] = itemValue.Value;
                }
            }
            listItem.Update();
            web.Context.Load(listItem);
            web.Context.ExecuteQueryRetry();
            var itemId = listItem.Id;
            foreach (var itemValue in updateValues.Where(u => u.FieldTypeString == "TaxonomyFieldTypeMulti" || u.FieldTypeString == "TaxonomyFieldType"))
            {
                switch (itemValue.FieldTypeString)
                {
                    case "TaxonomyFieldTypeMulti":
                        {
                            var field = listFields.FirstOrDefault(f => f.InternalName == itemValue.Key || f.Title == itemValue.Key);
                            var taxField = web.Context.CastTo<TaxonomyField>(field);
                            if (itemValue.Value != null)
                            {
                                var valueCollection = new TaxonomyFieldValueCollection(web.Context, string.Join(";#", itemValue.Value as IEnumerable<string>), taxField);
                                taxField.SetFieldValueByValueCollection(listItem, valueCollection);
                            }
                            else
                            {
                                var valueCollection = new TaxonomyFieldValueCollection(web.Context, null, taxField);
                                taxField.SetFieldValueByValueCollection(listItem, valueCollection);
                            }
                            listItem.Update();
                            web.Context.Load(listItem);
                            web.Context.ExecuteQueryRetry();
                            break;
                        }
                    case "TaxonomyFieldType":
                        {
                            var field = listFields.FirstOrDefault(f => f.InternalName == itemValue.Key || f.Title == itemValue.Key);
                            var taxField = web.Context.CastTo<TaxonomyField>(field);
                            taxField.EnsureProperty(f => f.TextField);
                            var taxValue = new TaxonomyFieldValue();
                            if (itemValue.Value != null)
                            {
                                var termString = ((List<string>)itemValue.Value)[0];
                                taxValue.Label = termString.Split(new string[] { ";#" }, StringSplitOptions.None)[1].Split(new char[] { '|' })[0];
                                taxValue.TermGuid = termString.Split(new string[] { ";#" }, StringSplitOptions.None)[1].Split(new char[] { '|' })[1];
                                taxValue.WssId = -1;
                                taxField.SetFieldValueByValue(listItem, taxValue);
                            }
                            else
                            {
                                taxValue.Label = string.Empty;
                                taxValue.TermGuid = "11111111-1111-1111-1111-111111111111";
                                taxValue.WssId = -1;
                                Field hiddenField = listFields.GetById(taxField.TextField);
                                listItem.Context.Load(hiddenField, tf => tf.InternalName);
                                listItem.Context.ExecuteQueryRetry();
                                taxField.SetFieldValueByValue(listItem, taxValue); // this order of updates is important.
                                listItem[hiddenField.InternalName] = string.Empty; // this order of updates is important.
                            }
                            listItem.Update();
                            web.Context.Load(listItem);
                            web.Context.ExecuteQueryRetry();
                            break;
                        }
                }
            }
        }

        /// <summary>
        /// This method is present to preserve backward compatibility with old file property format
        /// </summary>
        /// <typeparam name="T">Expected type</typeparam>
        /// <param name="jsonValue">json value</param>
        /// <param name="result">The result object, if success</param>
        /// <returns>Returns <c>true</c> if the value was sucessfully deserialized from the json string. Otherwise <c>false</c></returns>
        private static bool TryDeserializeAsJson<T>(string jsonValue, out T result)
        {
            try
            {
                result = JsonUtility.Deserialize<T>(jsonValue);
                return true;
            }
            catch (Newtonsoft.Json.JsonException)
            {
                result = default(T);
                return false;
            }
            // Other exception are not to be catched
        }
    }
}