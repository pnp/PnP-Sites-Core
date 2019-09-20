using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{

    public class SerializedListInstance
    {
        public string SchemaXml { get; set; }
    }

    public class ListInstanceValidator : ValidatorBase
    {
        private bool isNoScriptSite = false;

        #region construction        
        public ListInstanceValidator(Web web) : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:Lists/pnp:ListInstance";

            // Check if this is not a noscript site as we're not allowed to update some properties
            isNoScriptSite = web.IsNoScriptSite();
        }

        public ListInstanceValidator(ClientContext cc) : this(cc.Web)
        {
            this.cc = cc;
        }

        #endregion

        #region Validation logic
        public bool Validate(ListInstanceCollection sourceCollection, ListInstanceCollection targetCollection, TokenParser tokenParser)
        {
            // Convert object collections to XML 
            List<SerializedListInstance> sourceLists = new List<SerializedListInstance>();
            List<SerializedListInstance> targetLists = new List<SerializedListInstance>();

            foreach (ListInstance list in sourceCollection)
            {
                // don't add hidden lists since they're not exported again...
                if (!list.Hidden)
                {
                    ProvisioningTemplate pt = new ProvisioningTemplate();
                    pt.Lists.Add(list);

                    sourceLists.Add(new SerializedListInstance() { SchemaXml = ExtractElementXml(pt) });
                }
            }

            foreach (ListInstance list in targetCollection)
            {
                ProvisioningTemplate pt = new ProvisioningTemplate();
                pt.Lists.Add(list);

                targetLists.Add(new SerializedListInstance() { SchemaXml = ExtractElementXml(pt) });
            }

            // Use XML validation logic to compare source and target
            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            parserSettings.Add("SchemaXml", null);
            bool isListsMatch = ValidateObjectsXML(sourceLists, targetLists, "SchemaXml", new List<string> { "Title" }, tokenParser, parserSettings);

            // Use CustomAction validator to validate the custom actions on the list
            if (!isNoScriptSite)
            {
                var sourceListsWithUserCustomActions = sourceCollection.Where(p => p.UserCustomActions.Any() == true);
                foreach (ListInstance list in sourceListsWithUserCustomActions)
                {
                    var targetList = targetCollection.Where(p => p.Title == list.Title).FirstOrDefault();
                    if (!CustomActionValidator.ValidateCustomActions(list.UserCustomActions, targetList.UserCustomActions, tokenParser))
                    {
                        isListsMatch = false;
                        break;
                    }
                }
            }

            Console.WriteLine("-- Lists validation " + isListsMatch);
            return isListsMatch;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            #region Property handling
            // Base property handling
            // Drop list properties if they're not provided in the source XML
            string[] ListProperties = new string[] { "Description", "DocumentTemplate", "MinorVersionLimit", "MaxVersionLimit", "DraftVersionVisibility", "TemplateFeatureID",
                "EnableAttachments", "DefaultDisplayFormUrl", "DefaultEditFormUrl", "DefaultNewFormUrl", "ImageUrl", "ValidationFormula", "ValidationMessage" };
            foreach (string listProperty in ListProperties)
            {
                if (sourceObject.Attribute(listProperty) == null)
                {
                    DropAttribute(targetObject, listProperty);
                }
            }

            // Drop list elements if they're not provided in the source XML
            string[] ListElements = new string[] {
                "ContentTypeBindings", "Views", "FieldRefs", "Fields"
#if ONPREMISES
                , "Webhooks"
#endif
            };
            foreach (var listElement in ListElements)
            {
                var sourceListElementXML = sourceObject.Descendants(ns + listElement);
                var targetListElementXML = targetObject.Descendants(ns + listElement);
                if (sourceListElementXML.Any() == false && targetListElementXML.Any() == true)
                {
                    targetListElementXML.Remove();
                }

#if ONPREMISES
                // Drop webhooks element from on-premises validation flow
                if (listElement.Equals("Webhooks", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (sourceListElementXML.Any())
                    {
                        sourceListElementXML.Remove();
                    }
                }
#endif
            }

            // Drop TemplateFeatureID from both -  this was a temp measure, should not be needed anymore with fixed serializer
            //DropAttribute(targetObject, "TemplateFeatureID");
            //DropAttribute(sourceObject, "TemplateFeatureID");

#if ONPREMISES
            // MaxVersionLimit cannot be set in on-premises, so remove it before comparing before and after
            if (sourceObject.Attribute("MaxVersionLimit") != null)
            {
                DropAttribute(targetObject, "MaxVersionLimit");
                DropAttribute(sourceObject, "MaxVersionLimit");
                DropAttribute(targetObject, "ReadSecurity");
                DropAttribute(sourceObject, "ReadSecurity");
            }
#endif
            #endregion

            #region Contenttype handling
            // If RemoveExistingContentTypes is set then remove the attribute from source since on target we don't add this. 
            var contentTypeBindings = targetObject.Descendants(ns + "ContentTypeBinding");
            bool removeExistingContentTypes = false;
            if (sourceObject.Attribute("RemoveExistingContentTypes") != null)
            {
                removeExistingContentTypes = sourceObject.Attribute("RemoveExistingContentTypes").Value.ToBoolean();
                DropAttribute(sourceObject, "RemoveExistingContentTypes");
            }

            if (contentTypeBindings != null && contentTypeBindings.Any())
            {
                // One can add ContentTypeBindings without specifying ContentTypesEnabled. The engine will turn on ContentTypesEnabled automatically in that case
                if (sourceObject.Attribute("ContentTypesEnabled") == null)
                {
                    DropAttribute(targetObject, "ContentTypesEnabled");
                }

                if (removeExistingContentTypes)
                {
                    foreach (var contentTypeBinding in contentTypeBindings.ToList())
                    {
                        // Remove the folder content type if present because we're not removing that one via RemoveExistingContentTypes
                        if (contentTypeBinding.Attribute("ContentTypeID").Value == "0x0120")
                        {
                            contentTypeBinding.Remove();
                        }
                    }
                }
                else // We did not remove existing content types
                {
                    var sourceContentTypeBindings = sourceObject.Descendants(ns + "ContentTypeBinding");
                    foreach (var contentTypeBinding in contentTypeBindings.ToList())
                    {
                        string value = contentTypeBinding.Attribute("ContentTypeID").Value;
                        // drop all content types which are not mentioned in the source
                        if (sourceContentTypeBindings.Where(p => p.Attribute("ContentTypeID").Value == value).Any() == false)
                        {
                            contentTypeBinding.Remove();
                        }
                    }
                }
            }
            #endregion

            #region FieldRef handling
            var fieldRefs = sourceObject.Descendants(ns + "FieldRef");
            if (fieldRefs != null && fieldRefs.Any())
            {
                foreach (var fieldRef in fieldRefs.ToList())
                {
                    // Drop the name attribute from the source fieldrefs since this is just an informational attribute
                    if (fieldRef.Attribute("Name") != null)
                    {
                        DropAttribute(fieldRef, "Name");
                    }

                    // Drop hidden fieldrefs since they're not retreived by the engine again
                    if (fieldRef.Attribute("Hidden") != null && fieldRef.Attribute("Hidden").Value.ToBoolean() == true)
                    {
                        fieldRef.Remove();
                    }
                }
            }
            var targetFieldRefs = targetObject.Descendants(ns + "FieldRef");
            if (targetFieldRefs != null && targetFieldRefs.Any())
            {
                foreach (var targetFieldRef in targetFieldRefs.ToList())
                {
                    // Drop the name attribute from the source fieldrefs since this is just an informational attribute
                    if (targetFieldRef.Attribute("Name") != null)
                    {
                        DropAttribute(targetFieldRef, "Name");
                    }

                    // Drop the fields which were not present in source (there's always some OOB fieldrefs)
                    if (!fieldRefs.Where(p => p.Attribute("ID").Value == targetFieldRef.Attribute("ID").Value).Any())
                    {
                        targetFieldRef.Remove();
                    }
                    else
                    {
                        // If the source object does not have a DisplayName attribute then also remove it from the target
                        if (!(fieldRefs.Where(p => p.Attribute("ID").Value == targetFieldRef.Attribute("ID").Value).FirstOrDefault().Attribute("DisplayName") != null))
                        {
                            DropAttribute(targetFieldRef, "DisplayName");
                        }
                    }
                }
            }
            #endregion

            #region Field handling
            var targetFields = targetObject.Descendants("Field");
            var sourceFields = sourceObject.Descendants("Field");

            if (sourceFields != null && sourceFields.Any())
            {
                foreach (var sourceField in sourceFields.ToList())
                {
                    // Ensure both target and source using the same casing
                    UpperCaseAttribute(sourceField, "ID");
                }
            }

            if (targetFields != null && targetFields.Any())
            {
                foreach (var targetField in targetFields.ToList())
                {
                    // Ensure both target and source using the same casing
                    UpperCaseAttribute(targetField, "ID");

                    // Drop attributes before comparison
                    DropAttribute(targetField, "SourceID");
                    DropAttribute(targetField, "StaticName");
                    DropAttribute(targetField, "ColName");
                    DropAttribute(targetField, "RowOrdinal");
                    DropAttribute(targetField, "Version");

                    // If target field does not exist in source then drop it (SPO can add additional fields e.g. _IsRecord field refering the _ComplianceFlags field)
                    if (sourceFields != null && sourceFields.Any())
                    {
                        if (!sourceFields.Where(p => p.Attribute("ID").Value.Equals(targetField.Attribute("ID").Value)).Any())
                        {
                            targetField.Remove();
                        }
                    }
                }
            }
            #endregion

            #region View handling
            var sourceViews = sourceObject.Descendants("View");
            var targetViews = targetObject.Descendants("View");

            if (sourceViews != null && sourceViews.Any())
            {
                int sourceViewCount = 0;
                int targetViewCount = 0;
                foreach (var sourceView in sourceViews)
                {
                    sourceViewCount++;
                    string sourceViewName = sourceView.Attribute("DisplayName").Value;

                    if (targetViews.Where(v => v.Attribute("DisplayName").Value.Equals(sourceViewName, StringComparison.InvariantCultureIgnoreCase)).First() != null)
                    {
                        targetViewCount++;
                    }
                }

                if (sourceViewCount == targetViewCount)
                {
                    // if RemoveExistingViews was checked then we should have the same count of source and target views
                    if (sourceObject.Descendants(ns + "Views").First().Attribute("RemoveExistingViews") != null && sourceObject.Descendants(ns + "Views").First().Attribute("RemoveExistingViews").Value.ToBoolean() == true)
                    {
                        if (sourceViews.Count() == targetViews.Count())
                        {
                            // we've found the source views in the target + the original views were dropped, so we're good. Drop the view element to ensure valid XML comparison
                            sourceObject.Descendants(ns + "Views").Remove();
                            targetObject.Descendants(ns + "Views").Remove();
                        }
                    }
                    else
                    {
                        // we've found the source views in the target so we're good. Drop the view element to ensure valid XML comparison
                        sourceObject.Descendants(ns + "Views").Remove();
                        targetObject.Descendants(ns + "Views").Remove();
                    }
                }
            }
            #endregion

            #region FieldDefaults handling
            var sourceFieldDefaults = sourceObject.Descendants(ns + "FieldDefault");
            if (sourceFieldDefaults != null && sourceFieldDefaults.Any())
            {
                bool validFieldDefaults = true;

                foreach (var sourceFieldDefault in sourceFieldDefaults)
                {
                    string fieldDefaultValue = sourceFieldDefault.Value;
                    string fieldDefaultName = sourceFieldDefault.Attribute("FieldName").Value;

                    if (fieldDefaultValue != null && fieldDefaultName != null)
                    {
                        var targetField = targetFields.Where(p => p.Attribute("Name").Value.Equals(fieldDefaultName, StringComparison.InvariantCultureIgnoreCase)).First();
                        if (targetField.Descendants("Default").Any())
                        {
                            string targetFieldDefaultValue = targetField.Descendants("Default").First().Value;
                            if (!targetFieldDefaultValue.Equals(fieldDefaultValue, StringComparison.InvariantCultureIgnoreCase))
                            {
                                validFieldDefaults = false;
                            }
                            else
                            {
                                // remove the Default node
                                targetField.Descendants("Default").First().Remove();
                            }
                        }
                        else
                        {
                            validFieldDefaults = false;
                        }
                    }
                }

                if (validFieldDefaults)
                {
                    sourceObject.Descendants(ns + "FieldDefaults").Remove();
                }
            }

            // Drop any remaining Default node
            foreach (var targetField in targetFields)
            {
                if (targetField.Descendants("Default").Any())
                {
                    targetField.Descendants("Default").First().Remove();
                }
            }
            #endregion

            #region Folder handling
            // Folders are not extracted, so manual validation needed
            var sourceFolders = sourceObject.Descendants(ns + "Folders");
            if (sourceFolders != null && sourceFolders.Any())
            {
                // use CSOM to verify the folders are listed in target and only remove the source folder from the XML when this is the case
                if (this.cc != null)
                {
                    var foldersValid = true;
                    var list = this.cc.Web.GetListByUrl(sourceObject.Attribute("Url").Value);
                    if (list != null)
                    {
                        list.EnsureProperty(w => w.RootFolder);
                        foreach (var folder in sourceFolders.Descendants(ns + "Folder"))
                        {
                            // only verify first level folders
                            if (folder.Parent.Equals(sourceFolders.First()))
                            {
                                if (!list.RootFolder.FolderExists(folder.Attribute("Name").Value))
                                {
                                    foldersValid = false;
                                }

                                // if the folder has a security element then verify this as well
                                var sourceFolderSecurity = folder.Descendants(ns + "Security");
                                if (sourceFolderSecurity != null && sourceFolderSecurity.Any())
                                {
                                    // convert XML in ObjectSecurity object
                                    ObjectSecurity sourceFolderSecurityElement = new ObjectSecurity();
                                    sourceFolderSecurityElement.ClearSubscopes = sourceFolderSecurity.Descendants(ns + "BreakRoleInheritance").First().Attribute("ClearSubscopes").Value.ToBoolean();
                                    sourceFolderSecurityElement.CopyRoleAssignments = sourceFolderSecurity.Descendants(ns + "BreakRoleInheritance").First().Attribute("CopyRoleAssignments").Value.ToBoolean();

                                    var sourceRoleAssignments = folder.Descendants(ns + "RoleAssignment");
                                    foreach (var sourceRoleAssignment in sourceRoleAssignments)
                                    {
                                        sourceFolderSecurityElement.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment()
                                        {
                                            Principal = sourceRoleAssignment.Attribute("Principal").Value,
                                            RoleDefinition = sourceRoleAssignment.Attribute("RoleDefinition").Value
                                        });
                                    }

                                    // grab the "Securable" for the folder
                                    var currentFolderItem = list.RootFolder.EnsureFolder(folder.Attribute("Name").Value).ListItemAllFields;
                                    cc.Load(currentFolderItem);
                                    cc.ExecuteQueryRetry();

                                    // use CSOM to verify security settings
                                    if (ValidateSecurityCSOM(this.cc, sourceFolderSecurityElement, currentFolderItem))
                                    {
                                        sourceFolderSecurity.Remove();
                                    }
                                }
                            }
                        }
                    }

                    if (foldersValid)
                    {
                        sourceFolders.Remove();
                    }
                }
            }
            #endregion

            #region DataRows handling
            var sourceDataRows = sourceObject.Descendants(ns + "DataRow");
            if (sourceDataRows != null && sourceDataRows.Any())
            {
                bool dataRowsValidated = true;

                var list = this.cc.Web.GetListByUrl(sourceObject.Attribute("Url").Value);
                if (list != null)
                {
                    int dataRowCount = 0;
                    foreach (var sourceDataRow in sourceDataRows)
                    {
                        // Convert XML in DataRow object
                        DataRow sourceDataRowElement = null;
                        Dictionary<string, string> values = new Dictionary<string, string>();
                        foreach (var dataValue in sourceDataRow.Descendants(ns + "DataValue"))
                        {
                            values.Add(dataValue.Attribute("FieldName").Value, dataValue.Value);
                        }

                        ObjectSecurity sourceDataRowSecurityElement = null;
                        var sourceDataRowSecurity = sourceDataRow.Descendants(ns + "Security");
                        if (sourceDataRowSecurity != null && sourceDataRowSecurity.Any())
                        {
                            // convert XML in ObjectSecurity object
                            sourceDataRowSecurityElement = new ObjectSecurity();
                            sourceDataRowSecurityElement.ClearSubscopes = sourceDataRowSecurity.Descendants(ns + "BreakRoleInheritance").First().Attribute("ClearSubscopes").Value.ToBoolean();
                            sourceDataRowSecurityElement.CopyRoleAssignments = sourceDataRowSecurity.Descendants(ns + "BreakRoleInheritance").First().Attribute("CopyRoleAssignments").Value.ToBoolean();

                            var sourceRoleAssignments = sourceDataRowSecurity.Descendants(ns + "RoleAssignment");
                            foreach (var sourceRoleAssignment in sourceRoleAssignments)
                            {
                                sourceDataRowSecurityElement.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment()
                                {
                                    Principal = sourceRoleAssignment.Attribute("Principal").Value,
                                    RoleDefinition = sourceRoleAssignment.Attribute("RoleDefinition").Value
                                });
                            }
                        }

                        if (sourceDataRowSecurityElement != null)
                        {
                            sourceDataRowElement = new DataRow(values, sourceDataRowSecurityElement);
                        }
                        else
                        {
                            sourceDataRowElement = new DataRow(values);
                        }

                        dataRowCount++;
                        ListItem itemToValidate = null;
                        try
                        {
                            itemToValidate = list.GetItemById(dataRowCount);
                        }
                        catch
                        { }

                        if (itemToValidate == null || !ValidateDataRowsCSOM(cc, sourceDataRowElement, itemToValidate))
                        {
                            dataRowsValidated = false;
                        }
                    }
                }

                // If all datarows are validated then we can drop 
                if (dataRowsValidated)
                {
                    sourceObject.Descendants(ns + "DataRows").First().Remove();
                }
            }
            #endregion

            #region Security handling
            var sourceSecurity = sourceObject.Descendants(ns + "Security");
            if (sourceSecurity != null && sourceSecurity.Any())
            {
                var targetSecurity = targetObject.Descendants(ns + "Security");
                if (ValidateSecurityXml(sourceSecurity.First(), targetSecurity.First()))
                {
                    sourceSecurity.Remove();
                    targetSecurity.Remove();
                }
            }
            #endregion

            #region CustomAction handling
            var sourceCustomActions = sourceObject.Descendants(ns + "UserCustomActions");
            if (sourceCustomActions != null && sourceCustomActions.Any())
            {
                // delete custom actions since we validate these latter on
                var targetCustomActions = targetObject.Descendants(ns + "UserCustomActions");

                sourceCustomActions.Remove();
                if (targetCustomActions != null && targetCustomActions.Any())
                {
                    targetCustomActions.Remove();
                }
            }
            #endregion
        }

#endregion
    }
}
