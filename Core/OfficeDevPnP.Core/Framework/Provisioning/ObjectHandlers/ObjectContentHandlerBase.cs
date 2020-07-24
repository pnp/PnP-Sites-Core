using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal abstract class ObjectContentHandlerBase : ObjectHandlerBase
    {
        internal Model.File RetrieveFieldValues(Web web, Microsoft.SharePoint.Client.File file, Model.File modelFile)
        {
            ListItem listItem = null;
            try
            {
                listItem = file.EnsureProperty(f => f.ListItemAllFields);
            }
            catch { }

            if (listItem != null)
            {
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
                    "AccessPolicy",
                    "BSN",
                    "_ListSchemaVersion",
                    "_Dirty",
                    "_Parsable",
                    "_StubFile",
                    "_VirusStatus",
                    "_VirusVendorID",
                    "_CheckinComment"
                };

                foreach (var fieldValue in fieldValues.Where(f => !fieldstoExclude.Contains(f.Key)))
                {
                    if (fieldValue.Value != null && !string.IsNullOrEmpty(fieldValue.Value.ToString()))
                    {
                        var field = fields.FirstOrDefault(fs => fs.InternalName == fieldValue.Key);

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
#if !ONPREMISES
                                    value = userFieldValue.Email;
#else
                                value = userFieldValue.LookupValue;
#endif
                                }
                                break;
                            case "LookupMulti":
                                var lookupFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldLookupValue[];
                                if (lookupFieldValue != null)
                                {
                                    value = Tokenize(JsonUtility.Serialize(lookupFieldValue), web.Url);
                                }
                                break;
                            case "TaxonomyFieldType":
                                var taxonomyFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue;
                                if (taxonomyFieldValue != null)
                                {
                                    value = Tokenize(JsonUtility.Serialize(taxonomyFieldValue), web.Url);
                                }
                                break;
                            case "TaxonomyFieldTypeMulti":
                                var taxonomyMultiFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValueCollection;
                                if (taxonomyMultiFieldValue != null)
                                {
                                    value = Tokenize(JsonUtility.Serialize(taxonomyMultiFieldValue), web.Url);
                                }
                                break;
                            case "ContentTypeIdFieldType":
                            default:
                                value = Tokenize(fieldValue.Value.ToString(), web.Url, web);
                                break;
                        }

                        if (fieldValue.Key == "ContentTypeId")
                        {
                            // Replace the content typeid with a token
                            var ct = list.GetContentTypeById(value);
                            if (ct != null)
                            {
                                value = $"{{contenttypeid:{ct.Name}}}";
                            }
                        }

                        // We process real values only
                        if (value != null && !String.IsNullOrEmpty(value) && value != "[]")
                        {
                            modelFile.Properties.Add(fieldValue.Key, value);
                        }
                    }
                }
            }

            return modelFile;
        }

        internal void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string folderPath, string fileName, Boolean decodeFileName = false)
        {
            if (creationInfo.FileConnector != null)
            {
                var fileConnector = creationInfo.FileConnector;
                SharePointConnector connector = new SharePointConnector(web.Context, web.Url, "dummy");
                Uri u = new Uri(web.Url);

                if (u.PathAndQuery != "/")
                {
                    if (folderPath.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1)
                    {
                        folderPath = folderPath.Replace(u.PathAndQuery, "");
                    }
                }

                String container = folderPath.Trim('/').Replace("%20", " ").Replace("/", "\\");
                String persistenceFileName = (decodeFileName ? HttpUtility.UrlDecode(fileName) : fileName).Replace("%20", " ");

                if (fileConnector.Parameters.ContainsKey(FileConnectorBase.CONTAINER))
                {
                    container = string.Concat(fileConnector.GetContainer(), container);
                }

                using (Stream s = connector.GetFileStream(fileName, folderPath))
                {
                    if (s != null)
                    {
                        creationInfo.FileConnector.SaveFileStream(
                            persistenceFileName, container, s);
                    }
                }
            }
            else
            {
                WriteMessage($"No connector present to persist {fileName}.", ProvisioningMessageType.Error);
                scope.LogError($"No connector present to persist {fileName}.");
            }
        }
    }
}
