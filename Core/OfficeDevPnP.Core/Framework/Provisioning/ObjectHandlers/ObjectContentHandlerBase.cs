using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
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
                "_ModerationStatus",
                "HtmlDesignAssociated",
                "HtmlDesignStatusAndPreview",
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
                    else
                    {
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
                    modelFile.Properties.Add(fieldValue.Key, value);
                }
            }

            return modelFile;
        }

        internal void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string folderPath, string fileName, Boolean decodeFileName = false)
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
                        creationInfo.FileConnector.SaveFileStream(decodeFileName ? HttpUtility.UrlDecode(fileName) : fileName, s);
                    }
                }
            }
            else
            {
                WriteWarning("No connector present to persist homepage.", ProvisioningMessageType.Error);
                scope.LogError("No connector present to persist homepage");
            }
        }
    }
}
