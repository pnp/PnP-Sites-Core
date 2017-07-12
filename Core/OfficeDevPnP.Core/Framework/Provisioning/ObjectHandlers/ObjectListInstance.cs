using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using Field = Microsoft.SharePoint.Client.Field;
using View = OfficeDevPnP.Core.Framework.Provisioning.Model.View;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Text.RegularExpressions;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client.WebParts;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectListInstance : ObjectHandlerBase
    {

        public override string Name
        {
            get { return "List instances"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Lists.Any())
                {
                    var rootWeb = ((ClientContext)web.Context).Site.RootWeb;

                    web.EnsureProperties(w => w.ServerRelativeUrl, w => w.SupportedUILanguageIds);

                    web.Context.Load(web.Lists, lc => lc.IncludeWithDefaultProperties(l => l.RootFolder.ServerRelativeUrl));
                    web.Context.ExecuteQueryRetry();
                    var existingLists = web.Lists.AsEnumerable().ToList();
                    var serverRelativeUrl = web.ServerRelativeUrl;

                    var processedLists = new List<ListInfo>();

                    // Check if this is not a noscript site as we're not allowed to update some properties
                    bool isNoScriptSite = web.IsNoScriptSite();

                    var total = template.Lists.Count;

                    #region Lists

                    var currentListIndex = 0;
                    foreach (var templateList in template.Lists)
                    {
                        currentListIndex++;
                        WriteMessage($"List|{templateList.Title}|{currentListIndex}|{total}", ProvisioningMessageType.Progress);
                        // Check for the presence of the references content types and throw an exception if not present or in template
                        if (templateList.ContentTypesEnabled)
                        {
                            var existingCts = web.Context.LoadQuery(web.AvailableContentTypes);
                            web.Context.ExecuteQueryRetry();
                            foreach (var ct in templateList.ContentTypeBindings)
                            {
                                var found = template.ContentTypes.Any(t => t.Id.ToUpperInvariant() == ct.ContentTypeId.ToUpperInvariant());
                                if (found == false)
                                {
                                    found = existingCts.Any(t => t.StringId.ToUpperInvariant() == ct.ContentTypeId.ToUpperInvariant());
                                }
                                if (!found)
                                {
                                    scope.LogError("Referenced content type {0} not available in site or in template", ct.ContentTypeId);
                                    throw new Exception($"Referenced content type {ct.ContentTypeId} not available in site or in template");
                                }
                            }
                        }
                        // check if the List exists by url or by title
                        var index = existingLists.FindIndex(x => x.Title.Equals(parser.ParseString(templateList.Title), StringComparison.OrdinalIgnoreCase) || x.RootFolder.ServerRelativeUrl.Equals(UrlUtility.Combine(serverRelativeUrl, parser.ParseString(templateList.Url)), StringComparison.OrdinalIgnoreCase));

                        if (index == -1)
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_list__0_, templateList.Title);
                                var returnTuple = CreateList(web, templateList, parser, scope, isNoScriptSite);
                                var createdList = returnTuple.Item1;
                                parser = returnTuple.Item2;
                                processedLists.Add(new ListInfo { SiteList = createdList, TemplateList = templateList });

                                parser.AddToken(new ListIdToken(web, createdList.Title, createdList.Id));

#if !SP2013
                                foreach (var supportedlanguageId in web.SupportedUILanguageIds)
                                {
                                    var ci = new System.Globalization.CultureInfo(supportedlanguageId);
                                    var titleResource = createdList.TitleResource.GetValueForUICulture(ci.Name);
                                    createdList.Context.ExecuteQueryRetry();

                                    if (titleResource != null && titleResource.Value != null)
                                        parser.AddToken(new ListIdToken(web, titleResource.Value, createdList.Id));
                                }
#endif
                                parser.AddToken(new ListUrlToken(web, createdList.Title, createdList.RootFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length + 1)));
                            }
                            catch (Exception ex)
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_list__0__failed___1_____2_, templateList.Title, ex.Message, ex.StackTrace);
                                throw;
                            }
                        }
                        else
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Updating_list__0_, templateList.Title);
                                var existingList = web.Lists[index];
                                var returnTuple = UpdateList(web, existingList, templateList, parser, scope, isNoScriptSite);
                                var updatedList = returnTuple.Item1;
                                parser = returnTuple.Item2;
                                if (updatedList != null)
                                {
                                    processedLists.Add(new ListInfo { SiteList = updatedList, TemplateList = templateList });
                                }
                            }
                            catch (Exception ex)
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Updating_list__0__failed___1_____2_, templateList.Title, ex.Message, ex.StackTrace);
                                throw;
                            }
                        }
                    }
                    #endregion

                    #region FieldRefs

                    foreach (var listInfo in processedLists)
                    {

                        if (listInfo.TemplateList.FieldRefs.Any())
                        {
                            total = listInfo.TemplateList.FieldRefs.Count;
                            currentListIndex = 0;
                            foreach (var fieldRef in listInfo.TemplateList.FieldRefs)
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_FieldRef_Updating_list__0_, listInfo.TemplateList.Title, fieldRef.Name);

                                currentListIndex++;
                                WriteMessage($"Site Columns for list {listInfo.TemplateList.Title}|{fieldRef.Name}|{currentListIndex}|{total}", ProvisioningMessageType.Progress);
                                var field = rootWeb.GetFieldById(fieldRef.Id);
                                if (field == null)
                                {
                                    // log missing referenced field
                                    this.WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_InvalidFieldReference, listInfo.TemplateList.Title, fieldRef.Name, fieldRef.Id), ProvisioningMessageType.Error);

                                    // move onto next field reference
                                    continue;
                                }

                                if (!listInfo.SiteList.FieldExistsById(fieldRef.Id))
                                {
                                    field = CreateFieldRef(listInfo, field, fieldRef, parser);
                                }
                                else
                                {
                                    field = UpdateFieldRef(listInfo.SiteList, field.Id, fieldRef, parser);
                                }

                                field.EnsureProperties(f => f.InternalName, f => f.Title);

                                parser.AddToken(new FieldTitleToken(web, field.InternalName, field.Title));

#if !SP2013
                                var siteField = template.SiteFields.FirstOrDefault(f => Guid.Parse(XElement.Parse(f.SchemaXml).Attribute("ID").Value).Equals(field.Id));

                                if (siteField != null && siteField.SchemaXml.ContainsResourceToken())
                                {
                                    var isDirty = false;
                                    var originalFieldElement = XElement.Parse(siteField.SchemaXml);
                                    var nameAttributeValue = originalFieldElement.Attribute("DisplayName") != null ? originalFieldElement.Attribute("DisplayName").Value : "";
                                    if (nameAttributeValue.ContainsResourceToken())
                                    {
                                        if (field.TitleResource.SetUserResourceValue(nameAttributeValue, parser))
                                        {
                                            isDirty = true;
                                        }
                                    }
                                    var descriptionAttributeValue = originalFieldElement.Attribute("Description") != null ? originalFieldElement.Attribute("Description").Value : "";
                                    if (descriptionAttributeValue.ContainsResourceToken())
                                    {
                                        if (field.DescriptionResource.SetUserResourceValue(descriptionAttributeValue, parser))
                                        {
                                            isDirty = true;
                                        }
                                    }

                                    if (isDirty)
                                    {
                                        field.Update();
                                        field.Context.ExecuteQueryRetry();
                                    }
                                }
#endif
                            }

                            listInfo.SiteList.Update();
                            web.Context.ExecuteQueryRetry();
                        }
                    }

                    #endregion

                    #region Fields

                    foreach (var listInfo in processedLists)
                    {
                        if (listInfo.TemplateList.Fields.Any())
                        {
                            var currentFieldIndex = 0;
                            total = listInfo.TemplateList.Fields.Count;
                            foreach (var field in listInfo.TemplateList.Fields)
                            {

                                var fieldElement = XElement.Parse(parser.ParseString(field.SchemaXml, "~sitecollection", "~site"));
                                if (fieldElement.Attribute("ID") == null)
                                {
                                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field_schema_has_no_ID_attribute___0_, field.SchemaXml);
                                    throw new Exception(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field_schema_has_no_ID_attribute___0_, field.SchemaXml));
                                }
                                var id = fieldElement.Attribute("ID").Value;
                                var internalName = fieldElement.Attribute("InternalName")?.Value;

                                WriteMessage($"List Columns for list {listInfo.TemplateList.Title}|{internalName ?? id}|{currentFieldIndex}|{total}", ProvisioningMessageType.Progress);
                                Guid fieldGuid;
                                if (!Guid.TryParse(id, out fieldGuid))
                                {
                                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_ID_for_field_is_not_a_valid_Guid___0_, field.SchemaXml);
                                    throw new Exception(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_ID_for_field_is_not_a_valid_Guid___0_, id));
                                }
                                else
                                {
                                    var fieldFromList = listInfo.SiteList.GetFieldById<Field>(fieldGuid);
                                    if (fieldFromList == null)
                                    {
                                        try
                                        {
                                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_field__0_, fieldGuid);
                                            var createdField = CreateField(fieldElement, listInfo, parser, field.SchemaXml, web.Context, scope);
                                            if (createdField != null)
                                            {
                                                createdField.EnsureProperties(f => f.InternalName, f => f.Title);
                                                parser.AddToken(new FieldTitleToken(web, createdField.InternalName,
                                                    createdField.Title));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_field__0__failed___1_____2_, fieldGuid, ex.Message, ex.StackTrace);
                                            throw;
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Updating_field__0_, fieldGuid);
                                            var updatedField = UpdateField(web, listInfo, fieldGuid, fieldElement, fieldFromList, scope, parser, field.SchemaXml);
                                            if (updatedField != null)
                                            {
                                                updatedField.EnsureProperties(f => f.InternalName, f => f.Title);
                                                parser.AddToken(new FieldTitleToken(web, updatedField.InternalName,
                                                    updatedField.Title));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Updating_field__0__failed___1_____2_, fieldGuid, ex.Message, ex.StackTrace);
                                            throw;
                                        }

                                    }
                                }
                            }
                            listInfo.SiteList.Update();
                            web.Context.ExecuteQueryRetry();
                        }
                    }

                    #endregion

                    #region Default Field Values
                    foreach (var listInfo in processedLists)
                    {
                        if (listInfo.TemplateList.FieldDefaults.Any())
                        {
                            foreach (var fieldDefault in listInfo.TemplateList.FieldDefaults)
                            {
                                var field = listInfo.SiteList.Fields.GetByInternalNameOrTitle(fieldDefault.Key);
                                field.DefaultValue = fieldDefault.Value;
                                field.Update();
                                web.Context.ExecuteQueryRetry();
                            }
                        }
                    }
                    #endregion

                    #region Views



                    foreach (var listInfo in processedLists)
                    {


                        var list = listInfo.TemplateList;
                        var createdList = listInfo.SiteList;

                        if (list.Views.Any() && list.RemoveExistingViews)
                        {
                            while (createdList.Views.Any())
                            {
                                createdList.Views[0].DeleteObject();
                            }
                            web.Context.ExecuteQueryRetry();
                        }

                        var existingViews = createdList.Views;
                        web.Context.Load(existingViews, vs => vs.Include(v => v.Title, v => v.Id));
                        web.Context.ExecuteQueryRetry();
                        total = list.Views.Count;
                        var currentViewIndex = 0;
                        foreach (var view in list.Views)
                        {
                            currentViewIndex++;
                            CreateView(web, view, existingViews, createdList, scope, parser, currentViewIndex, total);

                        }
                    }

                    #endregion

                    #region Folders

                    // Folders are supported for document libraries and generic lists only
                    foreach (var list in processedLists)
                    {
                        list.SiteList.EnsureProperties(l => l.BaseType);
                        if ((list.SiteList.BaseType == BaseType.DocumentLibrary |
                            list.SiteList.BaseType == BaseType.GenericList) &&
                            list.TemplateList.Folders != null && list.TemplateList.Folders.Count > 0)
                        {
                            list.SiteList.EnableFolderCreation = true;
                            list.SiteList.Update();
                            web.Context.ExecuteQueryRetry();

                            var rootFolder = list.SiteList.RootFolder;
                            foreach (var folder in list.TemplateList.Folders)
                            {
                                CreateFolderInList(rootFolder, folder, parser, scope);
                            }
                        }
                    }

                    #endregion

                    // If an existing view is updated, and the list is to be listed on the QuickLaunch, it is removed because the existing view will be deleted and recreated from scratch. 
                    foreach (var listInfo in processedLists)
                    {
                        listInfo.SiteList.OnQuickLaunch = listInfo.TemplateList.OnQuickLaunch;
                        listInfo.SiteList.Update();
                    }
                    web.Context.ExecuteQueryRetry();

                    WriteMessage("Done processing lists", ProvisioningMessageType.Completed);
                }
            }
            return parser;
        }

        private void CreateView(Web web, View view, Microsoft.SharePoint.Client.ViewCollection existingViews, List createdList, PnPMonitoredScope monitoredScope, TokenParser parser, int currentViewIndex, int total)
        {
            try
            {
                var viewElement = XElement.Parse(view.SchemaXml);
                var displayNameElement = viewElement.Attribute("DisplayName");
                if (displayNameElement == null)
                {
                    throw new ApplicationException("Invalid View element, missing a valid value for the attribute DisplayName.");
                }
                WriteMessage($"Views for list {createdList.Title}|{displayNameElement.Value}|{currentViewIndex}|{total}", ProvisioningMessageType.Progress);
                monitoredScope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_view__0_, displayNameElement.Value);


                var viewTitle = parser.ParseString(displayNameElement.Value);
                var existingView = existingViews.FirstOrDefault(v => v.Title == viewTitle);
                if (existingView != null)
                {
                    existingView.DeleteObject();
                    web.Context.ExecuteQueryRetry();
                }

                // Type
                var viewTypeString = viewElement.Attribute("Type") != null ? viewElement.Attribute("Type").Value : "None";
                viewTypeString = viewTypeString[0].ToString().ToUpper() + viewTypeString.Substring(1).ToLower();
                var viewType = (ViewType)Enum.Parse(typeof(ViewType), viewTypeString);

                // Fields
                string[] viewFields = null;
                var viewFieldsElement = viewElement.Descendants("ViewFields").FirstOrDefault();
                if (viewFieldsElement != null)
                {
                    viewFields = (from field in viewElement.Descendants("ViewFields").Descendants("FieldRef") select field.Attribute("Name").Value).ToArray();
                }

                // Default view
                var viewDefault = viewElement.Attribute("DefaultView") != null && bool.Parse(viewElement.Attribute("DefaultView").Value);

                // Row limit
                var viewPaged = true;
                uint viewRowLimit = 30;
                var rowLimitElement = viewElement.Descendants("RowLimit").FirstOrDefault();
                if (rowLimitElement != null)
                {
                    if (rowLimitElement.Attribute("Paged") != null)
                    {
                        viewPaged = bool.Parse(rowLimitElement.Attribute("Paged").Value);
                    }
                    viewRowLimit = uint.Parse(rowLimitElement.Value);
                }

                // Query
                var viewQuery = new StringBuilder();
                foreach (var queryElement in viewElement.Descendants("Query").Elements())
                {
                    viewQuery.Append(queryElement.ToString());
                }

                var viewCI = new ViewCreationInformation
                {
                    ViewFields = viewFields,
                    RowLimit = viewRowLimit,
                    Paged = viewPaged,
                    Title = viewTitle,
                    Query = viewQuery.ToString(),
                    ViewTypeKind = viewType,
                    PersonalView = false,
                    SetAsDefaultView = viewDefault,
                };

                // Allow to specify a custom view url. View url is taken from title, so we first set title to the view url value we need, 
                // create the view and then set title back to the original value
                var urlAttribute = viewElement.Attribute("Url");
                var urlHasValue = urlAttribute != null && !string.IsNullOrEmpty(urlAttribute.Value);
                if (urlHasValue)
                {
                    //set Title to be equal to url (in order to generate desired url)
                    viewCI.Title = Path.GetFileNameWithoutExtension(urlAttribute.Value);
                }

                var createdView = createdList.Views.Add(viewCI);
                createdView.EnsureProperties(v => v.Scope, v => v.JSLink, v => v.Title, v => v.Aggregations, v => v.MobileView, v => v.MobileDefaultView);
                web.Context.ExecuteQueryRetry();

                if (urlHasValue)
                {
                    //restore original title 
                    createdView.Title = viewTitle;
                    createdView.Update();
                }

                // ContentTypeID
                var contentTypeID = viewElement.Attribute("ContentTypeID") != null ? viewElement.Attribute("ContentTypeID").Value : null;
                if (!string.IsNullOrEmpty(contentTypeID) && (contentTypeID != BuiltInContentTypeId.System))
                {
                    ContentTypeId childContentTypeId = null;
                    if (contentTypeID == BuiltInContentTypeId.RootOfList)
                    {
                        var childContentType = web.GetContentTypeById(contentTypeID);
                        childContentTypeId = childContentType != null ? childContentType.Id : null;
                    }
                    else
                    {
                        childContentTypeId = createdList.ContentTypes.BestMatch(contentTypeID);
                    }
                    if (childContentTypeId != null)
                    {
                        createdView.ContentTypeId = childContentTypeId;
                        createdView.Update();
                    }
                }

                // Default for content type
                bool parsedDefaultViewForContentType;
                var defaultViewForContentType = viewElement.Attribute("DefaultViewForContentType") != null ? viewElement.Attribute("DefaultViewForContentType").Value : null;
                if (!string.IsNullOrEmpty(defaultViewForContentType) && bool.TryParse(defaultViewForContentType, out parsedDefaultViewForContentType))
                {
                    createdView.DefaultViewForContentType = parsedDefaultViewForContentType;
                    createdView.Update();
                }

                // Scope
                var scope = viewElement.Attribute("Scope") != null ? viewElement.Attribute("Scope").Value : null;
                ViewScope parsedScope = ViewScope.DefaultValue;
                if (!string.IsNullOrEmpty(scope) && Enum.TryParse<ViewScope>(scope, out parsedScope))
                {
                    createdView.Scope = parsedScope;
                    createdView.Update();
                }

                // MobileView
                var mobileView = viewElement.Attribute("MobileView") != null && bool.Parse(viewElement.Attribute("MobileView").Value);
                if (mobileView)
                {
                    createdView.MobileView = mobileView;
                    createdView.Update();
                }

                // MobileDefaultView
                var mobileDefaultView = viewElement.Attribute("MobileDefaultView") != null && bool.Parse(viewElement.Attribute("MobileDefaultView").Value);
                if (mobileDefaultView)
                {
                    createdView.MobileDefaultView = mobileDefaultView;
                    createdView.Update();
                }

                // Aggregations
                var aggregationsElement = viewElement.Descendants("Aggregations").FirstOrDefault();
                if (aggregationsElement != null)
                {
                    if (aggregationsElement.HasElements)
                    {
                        var fieldRefString = "";
                        var fieldRefs = aggregationsElement.Descendants("FieldRef");
                        foreach (var fieldRef in fieldRefs)
                        {
                            fieldRefString += fieldRef.ToString();
                        }
                        if (createdView.Aggregations != fieldRefString)
                        {
                            createdView.Aggregations = fieldRefString;
                            createdView.Update();
                        }
                    }
                }


                // JSLink
                var jslinkElement = viewElement.Descendants("JSLink").FirstOrDefault();
                if (jslinkElement != null)
                {
                    var jslink = jslinkElement.Value;
                    if (createdView.JSLink != jslink)
                    {
                        createdView.JSLink = jslink;
                        createdView.Update();

                        // Only push the JSLink value to the web part as it contains a / indicating it's a custom one. So we're not pushing the OOB ones like clienttemplates.js or hierarchytaskslist.js
                        // but do push custom ones down to th web part (e.g. ~sitecollection/Style Library/JSLink-Samples/ConfidentialDocuments.js)
                        if (jslink.Contains("/"))
                        {
                            createdView.EnsureProperty(v => v.ServerRelativeUrl);
                            createdList.SetJSLinkCustomizations(createdView.ServerRelativeUrl, jslink);
                        }
                    }
                }
                

                createdList.Update();
                web.Context.ExecuteQueryRetry();

                // Add ListViewId token parser
                createdView.EnsureProperty(v => v.Id);
                parser.AddToken(new ListViewIdToken(web, createdList.Title, createdView.Title, createdView.Id));

#if !SP2013
                // Localize view title
                if (displayNameElement.Value.ContainsResourceToken())
                {
                    createdView.LocalizeView(web, displayNameElement.Value, parser, monitoredScope);
                }
#endif
            }
            catch (Exception ex)
            {
                monitoredScope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_view_failed___0_____1_, ex.Message, ex.StackTrace);
                throw;
            }
        }

        private static Field UpdateFieldRef(List siteList, Guid fieldId, FieldRef fieldRef, TokenParser parser)
        {
            // find the field in the list
            var listField = siteList.Fields.GetById(fieldId);

            siteList.Context.Load(listField, f => f.Id, f => f.Title, f => f.Hidden, f => f.Required);
            siteList.Context.ExecuteQueryRetry();

            var isDirty = false;

#if !SP2013
            if (!string.IsNullOrEmpty(fieldRef.DisplayName) && (fieldRef.DisplayName != listField.Title || fieldRef.DisplayName.ContainsResourceToken()))
            {
                if (fieldRef.DisplayName.ContainsResourceToken())
                {
                    listField.TitleResource.SetUserResourceValue(fieldRef.DisplayName, parser);
                }
                else
                {
                    listField.Title = fieldRef.DisplayName;
                }
                isDirty = true;
            }
#else
            if (!string.IsNullOrEmpty(fieldRef.DisplayName) && fieldRef.DisplayName != listField.Title)
            {
                listField.Title = fieldRef.DisplayName;
                isDirty = true;
            }
#endif

            // We cannot configure Hidden property for Phonetic fields 
            if (!(siteList.BaseTemplate == (int)ListTemplateType.Contacts &&
                (fieldRef.Name.Equals("LastNamePhonetic", StringComparison.InvariantCultureIgnoreCase) ||
                fieldRef.Name.Equals("FirstNamePhonetic", StringComparison.InvariantCultureIgnoreCase) ||
                fieldRef.Name.Equals("CompanyPhonetic", StringComparison.InvariantCultureIgnoreCase))))
            {
                if (fieldRef.Hidden != listField.Hidden)
                {
                    listField.Hidden = fieldRef.Hidden;
                    isDirty = true;
                }
            }

            if (fieldRef.Required != listField.Required)
            {
                listField.Required = fieldRef.Required;
                isDirty = true;
            }

            if (isDirty)
            {
                listField.UpdateAndPushChanges(true);
                siteList.Context.ExecuteQueryRetry();
            }

            return listField;
        }

        private static Field CreateFieldRef(ListInfo listInfo, Field field, FieldRef fieldRef, TokenParser parser)
        {
            field.EnsureProperty(f => f.SchemaXmlWithResourceTokens);
            XElement element = XElement.Parse(field.SchemaXmlWithResourceTokens);

            element.SetAttributeValue("AllowDeletion", "TRUE");

            var calculatedField = field as FieldCalculated;
            if (calculatedField != null)
            {
                if (element.Element("Formula") != null)
                {
                    element.Element("Formula").Value = calculatedField.Formula;
                }
            }

            field.SchemaXml = element.ToString();

            //Field has column Validation
            if (element.Elements("Validation").FirstOrDefault() != null)
            {
                field.SchemaXml = ObjectField.TokenizeFieldValidationFormula(field, field.SchemaXml);
            }

            var createdField = listInfo.SiteList.Fields.Add(field);

            createdField.Context.Load(createdField, cf => cf.Id, cf => cf.Title, cf => cf.Hidden, cf => cf.Required);
            createdField.Context.ExecuteQueryRetry();

            var isDirty = false;

#if !SP2013
            if (!string.IsNullOrEmpty(fieldRef.DisplayName) && (createdField.Title != fieldRef.DisplayName || fieldRef.DisplayName.ContainsResourceToken()))
            {
                if (fieldRef.DisplayName.ContainsResourceToken())
                {
                    createdField.TitleResource.SetUserResourceValue(fieldRef.DisplayName, parser);
                }
                else
                {
                    createdField.Title = fieldRef.DisplayName;
                }
                isDirty = true;
            }
#else
            if (!string.IsNullOrEmpty(fieldRef.DisplayName) && createdField.Title != fieldRef.DisplayName)
            {
                createdField.Title = fieldRef.DisplayName;
                isDirty = true;
            }
#endif

            if (createdField.Hidden != fieldRef.Hidden)
            {
                createdField.Hidden = fieldRef.Hidden;
                isDirty = true;
            }
            if (createdField.Required != fieldRef.Required)
            {
                createdField.Required = fieldRef.Required;
                isDirty = true;
            }
            if (isDirty)
            {
                createdField.Update();
                createdField.Context.ExecuteQueryRetry();
            }

            return createdField;
        }

        private static Field CreateField(XElement fieldElement, ListInfo listInfo, TokenParser parser, string originalFieldXml, ClientRuntimeContext context, PnPMonitoredScope scope)
        {
            Field field = null;
            fieldElement = PrepareField(fieldElement);

            var fieldXml = parser.ParseString(fieldElement.ToString(), "~sitecollection", "~site");
            if (IsFieldXmlValid(parser.ParseString(originalFieldXml), parser, context))
            {
                field = listInfo.SiteList.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddFieldInternalNameHint);
                listInfo.SiteList.Context.Load(field);
                listInfo.SiteList.Context.ExecuteQueryRetry();

                bool isDirty = false;
#if !SP2013
                if (originalFieldXml.ContainsResourceToken())
                {
                    var originalFieldElement = XElement.Parse(originalFieldXml);
                    var nameAttributeValue = originalFieldElement.Attribute("DisplayName") != null ? originalFieldElement.Attribute("DisplayName").Value : "";
                    if (nameAttributeValue.ContainsResourceToken())
                    {
                        if (field.TitleResource.SetUserResourceValue(nameAttributeValue, parser))
                        {
                            isDirty = true;
                        }
                    }
                    var descriptionAttributeValue = originalFieldElement.Attribute("Description") != null ? originalFieldElement.Attribute("Description").Value : "";
                    if (descriptionAttributeValue.ContainsResourceToken())
                    {
                        if (field.DescriptionResource.SetUserResourceValue(descriptionAttributeValue, parser))
                        {
                            isDirty = true;
                        }
                    }
                }
#endif
                if (isDirty)
                {
                    field.Update();
                    listInfo.SiteList.Context.ExecuteQueryRetry();
                }
            }
            else
            {
                // The field Xml was found invalid
                var tokenString = parser.GetLeftOverTokens(originalFieldXml).Aggregate(String.Empty, (acc, i) => acc + " " + i);
                scope.LogError("The field was found invalid: {0}", tokenString);
                throw new Exception($"The field was found invalid: {tokenString}");
            }
            return field;
        }

        private Field UpdateField(ClientObject web, ListInfo listInfo, Guid fieldId, XElement templateFieldElement, Field existingField, PnPMonitoredScope scope, TokenParser parser, string originalFieldXml)
        {
            Field field = null;
            web.Context.Load(existingField, f => f.SchemaXmlWithResourceTokens);
            web.Context.ExecuteQueryRetry();

            var existingFieldElement = XElement.Parse(existingField.SchemaXmlWithResourceTokens);

            var equalityComparer = new XNodeEqualityComparer();

            // Is field different in template?
            if (equalityComparer.GetHashCode(existingFieldElement) != equalityComparer.GetHashCode(templateFieldElement))
            {
                // Is existing field of the same type?
                if (existingFieldElement.Attribute("Type").Value == templateFieldElement.Attribute("Type").Value)
                {
                    templateFieldElement = PrepareField(templateFieldElement);
                    if (IsFieldXmlValid(parser.ParseString(templateFieldElement.ToString()), parser, web.Context))
                    {
                        foreach (var attribute in templateFieldElement.Attributes())
                        {
                            if (existingFieldElement.Attribute(attribute.Name) != null)
                            {
                                existingFieldElement.Attribute(attribute.Name).Value = attribute.Value;
                            }
                            else
                            {
                                existingFieldElement.Add(attribute);
                            }
                        }
                        foreach (var element in templateFieldElement.Elements())
                        {
                            if (existingFieldElement.Element(element.Name) != null)
                            {
                                existingFieldElement.Element(element.Name).Remove();
                            }
                            existingFieldElement.Add(element);
                        }

                        if (string.Equals(templateFieldElement.Attribute("Type").Value, "Calculated", StringComparison.OrdinalIgnoreCase))
                        {
                            var fieldRefsElement = existingFieldElement.Descendants("FieldRefs").FirstOrDefault();
                            if (fieldRefsElement != null)
                            {
                                fieldRefsElement.Remove();
                            }
                        }

                        if (existingFieldElement.Attribute("Version") != null)
                        {
                            existingFieldElement.Attributes("Version").Remove();
                        }
                        existingField.SchemaXml = parser.ParseString(existingFieldElement.ToString(), "~sitecollection", "~site");
                        existingField.UpdateAndPushChanges(true);
                        web.Context.ExecuteQueryRetry();
                        bool isDirty = false;
#if !SP2013
                        if (originalFieldXml.ContainsResourceToken())
                        {
                            var originalFieldElement = XElement.Parse(originalFieldXml);
                            var nameAttributeValue = originalFieldElement.Attribute("DisplayName") != null ? originalFieldElement.Attribute("DisplayName").Value : "";
                            if (nameAttributeValue.ContainsResourceToken())
                            {
                                if (existingField.TitleResource.SetUserResourceValue(nameAttributeValue, parser))
                                {
                                    isDirty = true;
                                }
                            }
                            var descriptionAttributeValue = originalFieldElement.Attribute("Description") != null ? originalFieldElement.Attribute("Description").Value : "";
                            if (descriptionAttributeValue.ContainsResourceToken())
                            {
                                if (existingField.DescriptionResource.SetUserResourceValue(descriptionAttributeValue, parser))
                                {
                                    isDirty = true;
                                }
                            }
                        }
#endif
                        if (isDirty)
                        {
                            existingField.Update();
                            web.Context.ExecuteQueryRetry();
                            field = existingField;
                        }
                    }
                    else
                    {
                        // The field Xml was found invalid
                        var tokenString = parser.GetLeftOverTokens(originalFieldXml).Aggregate(String.Empty, (acc, i) => acc + " " + i);
                        scope.LogError("The field was found invalid: {0}", tokenString);
                        throw new Exception($"The field was found invalid: {tokenString}");
                    }
                }
                else
                {
                    var fieldName = existingFieldElement.Attribute("Name") != null ? existingFieldElement.Attribute("Name").Value : existingFieldElement.Attribute("StaticName").Value;
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field__0____1___exists_in_list__2____3___but_is_of_different_type__Skipping_field_, fieldName, fieldId, listInfo.TemplateList.Title, listInfo.SiteList.Id);
                    WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field__0____1___exists_in_list__2____3___but_is_of_different_type__Skipping_field_, fieldName, fieldId, listInfo.TemplateList.Title, listInfo.SiteList.Id), ProvisioningMessageType.Warning);
                }
            }
            return field;
        }

        private static XElement PrepareField(XElement fieldElement)
        {
            var listIdentifier = fieldElement.Attribute("List") != null ? fieldElement.Attribute("List").Value : null;

            if (listIdentifier != null)
            {
                // Temporary remove list attribute from fieldElement
                fieldElement.Attribute("List").Remove();

                if (fieldElement.Attribute("RelationshipDeleteBehavior") != null)
                {
                    if (fieldElement.Attribute("RelationshipDeleteBehavior").Value.Equals("Restrict") ||
                        fieldElement.Attribute("RelationshipDeleteBehavior").Value.Equals("Cascade"))
                    {
                        // If RelationshipDeleteBehavior is either 'Restrict' or 'Cascade',
                        // make sure that Indexed is set to TRUE
                        if (fieldElement.Attribute("Indexed") != null)
                            fieldElement.Attribute("Indexed").Value = "TRUE";
                        else
                            fieldElement.Add(new XAttribute("Indexed", "TRUE"));
                    }

                    fieldElement.Attribute("RelationshipDeleteBehavior").Remove();
                }
            }

            return fieldElement;
        }

        private Tuple<List, TokenParser> UpdateList(Web web, List existingList, ListInstance templateList, TokenParser parser, PnPMonitoredScope scope, bool isNoScriptSite = false)
        {
            web.Context.Load(existingList,
                l => l.Title,
                l => l.Description,
                l => l.OnQuickLaunch,
                l => l.Hidden,
                l => l.ContentTypesEnabled,
                l => l.EnableAttachments,
                l => l.EnableVersioning,
                l => l.EnableFolderCreation,
                l => l.EnableModeration,
                l => l.EnableMinorVersions,
                l => l.ForceCheckout,
                l => l.DraftVersionVisibility,
                l => l.Views,
                l => l.DocumentTemplateUrl,
                l => l.RootFolder,
                l => l.BaseType,
                l => l.BaseTemplate
#if !SP2013
, l => l.MajorWithMinorVersionsLimit
, l => l.MajorVersionLimit
#endif
);
            web.Context.ExecuteQueryRetry();

            if (existingList.BaseTemplate == templateList.TemplateType)
            {
                var isDirty = false;
                if (parser.ParseString(templateList.Title) != existingList.Title)
                {
                    var oldTitle = existingList.Title;
                    existingList.Title = parser.ParseString(templateList.Title);
                    if (!oldTitle.Equals(existingList.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        parser.AddToken(new ListIdToken(web, existingList.Title, existingList.Id));
                        parser.AddToken(new ListUrlToken(web, existingList.Title, existingList.RootFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length + 1)));
                    }
                    isDirty = true;
                }
                if (!string.IsNullOrEmpty(templateList.DocumentTemplate))
                {
                    if (existingList.DocumentTemplateUrl != parser.ParseString(templateList.DocumentTemplate))
                    {
                        existingList.DocumentTemplateUrl = parser.ParseString(templateList.DocumentTemplate);
                        isDirty = true;
                    }
                }
                if (!string.IsNullOrEmpty(templateList.Description) && templateList.Description != existingList.Description)
                {
                    existingList.Description = templateList.Description;
                    isDirty = true;
                }
                if (templateList.Hidden != existingList.Hidden)
                {
                    existingList.Hidden = templateList.Hidden;
                    isDirty = true;
                }
                if (templateList.OnQuickLaunch != existingList.OnQuickLaunch)
                {
                    existingList.OnQuickLaunch = templateList.OnQuickLaunch;
                    isDirty = true;
                }
                if (existingList.BaseTemplate != (int)ListTemplateType.Survey &&
                    templateList.ContentTypesEnabled != existingList.ContentTypesEnabled)
                {
                    existingList.ContentTypesEnabled = templateList.ContentTypesEnabled;
                    isDirty = true;
                }
                if (existingList.BaseTemplate != (int)ListTemplateType.Survey &&
                    existingList.BaseTemplate != (int)ListTemplateType.DocumentLibrary &&
                    existingList.BaseTemplate != (int)ListTemplateType.PictureLibrary &&
                    existingList.BaseTemplate != 850) // 850 = Pages library on publishing site
                {
                    // https://msdn.microsoft.com/EN-US/library/microsoft.sharepoint.splist.enableattachments.aspx
                    // The EnableAttachments property does not apply to any list that has a base type of Survey, DocumentLibrary or PictureLibrary.
                    // If you set this property to true for either type of list, it throws an SPException.
                    if (templateList.EnableAttachments != existingList.EnableAttachments)
                    {
                        existingList.EnableAttachments = templateList.EnableAttachments;
                        isDirty = true;
                    }
                }
                if (existingList.BaseTemplate != (int)ListTemplateType.DiscussionBoard)
                {
                    if (templateList.EnableFolderCreation != existingList.EnableFolderCreation)
                    {
                        existingList.EnableFolderCreation = templateList.EnableFolderCreation;
                        isDirty = true;
                    }
                }
#if !SP2013
                if (templateList.Title.ContainsResourceToken())
                {
                    if (existingList.TitleResource.SetUserResourceValue(templateList.Title, parser))
                    {
                        isDirty = true;
                    }
                }
#endif
                if (existingList.EnableModeration != templateList.EnableModeration)
                {
                    existingList.EnableModeration = templateList.EnableModeration;
                    isDirty = true;
                }

                if (templateList.ForceCheckout != existingList.ForceCheckout)
                {
                    existingList.ForceCheckout = templateList.ForceCheckout;
                    isDirty = true;
                }

                if (templateList.EnableVersioning)
                {
                    if (existingList.EnableVersioning != templateList.EnableVersioning)
                    {
                        existingList.EnableVersioning = templateList.EnableVersioning;
                        isDirty = true;
                    }
#if !SP2013
                    if (existingList.MajorVersionLimit != templateList.MaxVersionLimit)
                    {
                        existingList.MajorVersionLimit = templateList.MaxVersionLimit;
                        isDirty = true;
                    }
#endif
                    if (existingList.BaseType == BaseType.DocumentLibrary)
                    {
                        // Only supported on Document Libraries
                        if (templateList.EnableMinorVersions != existingList.EnableMinorVersions)
                        {
                            existingList.EnableMinorVersions = templateList.EnableMinorVersions;
                            isDirty = true;
                        }

                        if ((DraftVisibilityType)templateList.DraftVersionVisibility != existingList.DraftVersionVisibility)
                        {
                            existingList.DraftVersionVisibility = (DraftVisibilityType)templateList.DraftVersionVisibility;
                            isDirty = true;
                        }

                        if (templateList.EnableMinorVersions)
                        {
                            if (templateList.MinorVersionLimit != existingList.MajorWithMinorVersionsLimit)
                            {
                                existingList.MajorWithMinorVersionsLimit = templateList.MinorVersionLimit;
                            }

                            if (DraftVisibilityType.Approver ==
                                (DraftVisibilityType)templateList.DraftVersionVisibility)
                            {
                                if (templateList.EnableModeration)
                                {
                                    if ((DraftVisibilityType)templateList.DraftVersionVisibility != existingList.DraftVersionVisibility)
                                    {
                                        existingList.DraftVersionVisibility = (DraftVisibilityType)templateList.DraftVersionVisibility;
                                        isDirty = true;
                                    }
                                }
                            }
                            else
                            {
                                if ((DraftVisibilityType)templateList.DraftVersionVisibility != existingList.DraftVersionVisibility)
                                {
                                    existingList.DraftVersionVisibility = (DraftVisibilityType)templateList.DraftVersionVisibility;
                                    isDirty = true;
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (existingList.EnableVersioning != templateList.EnableVersioning)
                    {
                        existingList.EnableVersioning = templateList.EnableVersioning;
                        isDirty = true;
                    }
                }

                if (isDirty)
                {
                    existingList.Update();
                    web.Context.ExecuteQueryRetry();
                    isDirty = false;
                }

#region UserCustomActions
                if (!isNoScriptSite)
                {
                    // Add any UserCustomActions
                    var existingUserCustomActions = existingList.UserCustomActions;
                    web.Context.Load(existingUserCustomActions);
                    web.Context.ExecuteQueryRetry();

                    foreach (CustomAction userCustomAction in templateList.UserCustomActions)
                    {
                        // Check for existing custom actions before adding (compare by custom action name)
                        if (!existingUserCustomActions.AsEnumerable().Any(uca => uca.Name == userCustomAction.Name))
                        {
                            CreateListCustomAction(existingList, parser, userCustomAction);
                            isDirty = true;
                        }
                        else
                        {
                            var existingCustomAction = existingUserCustomActions.AsEnumerable().FirstOrDefault(uca => uca.Name == userCustomAction.Name);
                            if (existingCustomAction != null)
                            {
                                isDirty = true;

                                // If the custom action already exists
                                if (userCustomAction.Remove)
                                {
                                    // And if we need to remove it, we simply delete it
                                    existingCustomAction.DeleteObject();
                                }
                                else
                                {
                                    // Otherwise we update it, and before we force the target 
                                    // registration type and ID to avoid issues
                                    userCustomAction.RegistrationType = UserCustomActionRegistrationType.List;
                                    userCustomAction.RegistrationId = existingList.Id.ToString("B").ToUpper();
                                    ObjectCustomActions.UpdateCustomAction(parser, scope, userCustomAction, existingCustomAction);
                                    // Blank out these values again to avoid inconsistent domain model data
                                    userCustomAction.RegistrationType = UserCustomActionRegistrationType.None;
                                    userCustomAction.RegistrationId = null;
                                }
                            }
                        }
                    }

                    if (isDirty)
                    {
                        existingList.Update();
                        web.Context.ExecuteQueryRetry();
                        isDirty = false;
                    }
                }
                else
                {
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_SkipAddingOrUpdatingCustomActions);
                }
#endregion

                if (existingList.ContentTypesEnabled)
                {

                    // Check if we need to add a content type

                    var existingContentTypes = existingList.ContentTypes;
                    web.Context.Load(existingContentTypes, cts => cts.Include(ct => ct.StringId));
                    web.Context.ExecuteQueryRetry();

                    if (templateList.RemoveExistingContentTypes && existingContentTypes.Count > 0)
                    {
                        WriteMessage($"You specified to remove existing content types for the list  with url '{existingList.RootFolder.ServerRelativeUrl}'. We found a list with the same url in the site. In case of a list update we cannot remove existing content types as they can be in use by existing list items and/or documents.", ProvisioningMessageType.Warning);
                    }

                    var bindingsToAdd = templateList.ContentTypeBindings.Where(ctb => existingContentTypes.All(ct => !ctb.ContentTypeId.Equals(ct.StringId, StringComparison.InvariantCultureIgnoreCase))).ToList();
                    var defaultCtBinding = templateList.ContentTypeBindings.FirstOrDefault(ctb => ctb.Default == true);
                    var currentDefaultContentTypeId = existingContentTypes.First().StringId;

                    foreach (var ctb in bindingsToAdd)
                    {
                        var tempCT = web.GetContentTypeById(ctb.ContentTypeId, searchInSiteHierarchy: true);
                        if (tempCT != null)
                        {
                            // Get the name of the existing CT
                            var name = tempCT.EnsureProperty(ct => ct.Name);

                            // If the CT does not exist in the target list, and we don't have to remove it
                            if (!existingList.ContentTypeExistsByName(name) && !ctb.Remove)
                            {
                                existingList.AddContentTypeToListById(ctb.ContentTypeId, searchContentTypeInSiteHierarchy: true);
                            }
                            // Else if the CT exists in the target list, and we have to remove it
                            else if (existingList.ContentTypeExistsByName(name) && ctb.Remove)
                            {
                                // Then remove it from the target list
                                existingList.RemoveContentTypeByName(name);
                            }
                        }
                    }

                    // default ContentTypeBinding should be set last because 
                    // list extension .SetDefaultContentTypeToList() re-sets 
                    // the list.RootFolder UniqueContentTypeOrder property
                    // which may cause missing CTs from the "New Button"
                    if (defaultCtBinding != null)
                    {
                        // Only update the defualt contenttype when we detect a change in default value
                        if (!currentDefaultContentTypeId.Equals(defaultCtBinding.ContentTypeId, StringComparison.InvariantCultureIgnoreCase))
                        {
                            existingList.SetDefaultContentTypeToList(defaultCtBinding.ContentTypeId);
                        }
                    }
                }
                if (templateList.Security != null)
                {
                    existingList.SetSecurity(parser, templateList.Security);
                }
                return Tuple.Create(existingList, parser);
            }
            else
            {
                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_List__0____1____2___exists_but_is_of_a_different_type__Skipping_list_, templateList.Title, templateList.Url, existingList.Id);
                WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_List__0____1____2___exists_but_is_of_a_different_type__Skipping_list_, templateList.Title, templateList.Url, existingList.Id), ProvisioningMessageType.Warning);
                return null;
            }
        }

        private static void CreateListCustomAction(List existingList, TokenParser parser, CustomAction userCustomAction)
        {
            UserCustomAction newUserCustomAction = existingList.UserCustomActions.Add();

            newUserCustomAction.Title = userCustomAction.Title;
            newUserCustomAction.Description = userCustomAction.Description;

#if !ONPREMISES
            if (!string.IsNullOrEmpty(userCustomAction.Title) && userCustomAction.Title.ContainsResourceToken())
            {
                newUserCustomAction.TitleResource.SetUserResourceValue(userCustomAction.Title, parser);
            }
            if (!string.IsNullOrEmpty(userCustomAction.Description) && userCustomAction.Description.ContainsResourceToken())
            {
                newUserCustomAction.DescriptionResource.SetUserResourceValue(userCustomAction.Description, parser);
            }
#endif

            newUserCustomAction.Name = userCustomAction.Name;
            newUserCustomAction.ImageUrl = userCustomAction.ImageUrl;
            newUserCustomAction.Rights = userCustomAction.Rights;
            newUserCustomAction.Sequence = userCustomAction.Sequence;
            newUserCustomAction.Group = userCustomAction.Group;
            newUserCustomAction.Location = userCustomAction.Location;
            //newUserCustomAction.RegistrationId = userCustomAction.RegistrationId;
            //newUserCustomAction.RegistrationType = userCustomAction.RegistrationType;
            newUserCustomAction.CommandUIExtension =
                userCustomAction.CommandUIExtension != null ?
                    parser.ParseString(userCustomAction.CommandUIExtension.ToString()) :
                    string.Empty;
            newUserCustomAction.ScriptBlock = userCustomAction.ScriptBlock;
            newUserCustomAction.ScriptSrc = userCustomAction.ScriptSrc;
            newUserCustomAction.Url = userCustomAction.Url;

            newUserCustomAction.Update();
        }

        private Tuple<List, TokenParser> CreateList(Web web, ListInstance list, TokenParser parser, PnPMonitoredScope scope, bool isNoScriptSite = false)
        {
            List createdList;
            if (list.Url.Equals("SiteAssets") && list.TemplateType == (int)ListTemplateType.DocumentLibrary)
            {
                //Ensure that the Site Assets library is created using the out of the box creation mechanism
                //Site Assets that are created using the EnsureSiteAssetsLibrary method slightly differ from
                //default Document Libraries. See issue 512 (https://github.com/OfficeDev/PnP-Sites-Core/issues/512)
                //for details about the issue fixed by this approach.
                createdList = web.Lists.EnsureSiteAssetsLibrary();
                //Check that Title and Description have the correct values
                web.Context.Load(createdList, l => l.Title,
                                              l => l.Description);
                web.Context.ExecuteQueryRetry();
                var isDirty = false;
                if (!string.Equals(createdList.Description, list.Description))
                {
                    createdList.Description = list.Description;
                    isDirty = true;
                }
                if (!string.Equals(createdList.Title, list.Title))
                {
                    createdList.Title = list.Title;
                    isDirty = true;
                }
                if (isDirty)
                {
                    createdList.Update();
                    web.Context.ExecuteQueryRetry();
                }

            }
            else
            {
                var listCreate = new ListCreationInformation();
                listCreate.Description = list.Description;
                listCreate.TemplateType = list.TemplateType;
                listCreate.Title = parser.ParseString(list.Title);

                // the line of code below doesn't add the list to QuickLaunch
                // the OnQuickLaunch property is re-set on the Created List object
                listCreate.QuickLaunchOption = list.OnQuickLaunch ? QuickLaunchOptions.On : QuickLaunchOptions.Off;

                listCreate.Url = parser.ParseString(list.Url);
                listCreate.TemplateFeatureId = list.TemplateFeatureID;

                createdList = web.Lists.Add(listCreate);
                createdList.Update();
            }
            web.Context.Load(createdList, l => l.BaseTemplate);
            web.Context.ExecuteQueryRetry();

#if !SP2013
            if (list.Title.ContainsResourceToken())
            {
                createdList.TitleResource.SetUserResourceValue(list.Title, parser);
            }
            if (list.Description.ContainsResourceToken())
            {
                createdList.DescriptionResource.SetUserResourceValue(list.Description, parser);
            }
#endif
            if (!String.IsNullOrEmpty(list.DocumentTemplate))
            {
                createdList.DocumentTemplateUrl = parser.ParseString(list.DocumentTemplate);
            }

            // EnableAttachments are not supported for DocumentLibraries, Survey and PictureLibraries
            // TODO: the user should be warned
            if (createdList.BaseTemplate != (int)ListTemplateType.DocumentLibrary &&
                createdList.BaseTemplate != (int)ListTemplateType.Survey &&
                createdList.BaseTemplate != (int)ListTemplateType.PictureLibrary)
            {
                createdList.EnableAttachments = list.EnableAttachments;
            }

            createdList.EnableModeration = list.EnableModeration;
            createdList.ForceCheckout = list.ForceCheckout;

            // Done for all other lists than for Survey - With Surveys versioning configuration will cause an exception
            if (createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                createdList.EnableVersioning = list.EnableVersioning;
                if (list.EnableVersioning)
                {
#if !SP2013
                    createdList.MajorVersionLimit = list.MaxVersionLimit;
#endif
                    // DraftVisibilityType.Approver is available only when the EnableModeration option of the list is true
                    if (DraftVisibilityType.Approver ==
                        (DraftVisibilityType)list.DraftVersionVisibility)
                    {
                        if (list.EnableModeration)
                        {
                            createdList.DraftVersionVisibility =
                                (DraftVisibilityType)list.DraftVersionVisibility;
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_DraftVersionVisibility_not_applied_because_EnableModeration_is_not_set_to_true);
                            WriteMessage(CoreResources.Provisioning_ObjectHandlers_ListInstances_DraftVersionVisibility_not_applied_because_EnableModeration_is_not_set_to_true, ProvisioningMessageType.Warning);
                        }
                    }
                    else
                    {
                        createdList.DraftVersionVisibility = (DraftVisibilityType)list.DraftVersionVisibility;
                    }

                    if (createdList.BaseTemplate == (int)ListTemplateType.DocumentLibrary)
                    {
                        // Only supported on Document Libraries
                        createdList.EnableMinorVersions = list.EnableMinorVersions;
                        createdList.DraftVersionVisibility = (DraftVisibilityType)list.DraftVersionVisibility;

                        if (list.EnableMinorVersions)
                        {
                            createdList.MajorWithMinorVersionsLimit = list.MinorVersionLimit; // Set only if enabled, otherwise you'll get exception due setting value to zero.
                        }
                    }
                }
            }

            createdList.OnQuickLaunch = list.OnQuickLaunch;
            if (createdList.BaseTemplate != (int)ListTemplateType.DiscussionBoard &&
                createdList.BaseTemplate != (int)ListTemplateType.Events)
            {
                createdList.EnableFolderCreation = list.EnableFolderCreation;
            }
            createdList.Hidden = list.Hidden;

            if (createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                createdList.ContentTypesEnabled = list.ContentTypesEnabled;
            }

            createdList.Update();

            web.Context.Load(createdList.Views);
            web.Context.Load(createdList, l => l.Id);
            web.Context.Load(createdList, l => l.RootFolder.ServerRelativeUrl);
            web.Context.Load(createdList.ContentTypes);
            web.Context.ExecuteQueryRetry();


            if (createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                // Remove existing content types only if there are custom content type bindings
                var contentTypesToRemove = new List<ContentType>();
                if (list.RemoveExistingContentTypes && list.ContentTypeBindings.Count > 0)
                {
                    contentTypesToRemove.AddRange(createdList.ContentTypes);
                }

                ContentTypeBinding defaultCtBinding = null;
                foreach (var ctBinding in list.ContentTypeBindings)
                {
                    var tempCT = web.GetContentTypeById(ctBinding.ContentTypeId, searchInSiteHierarchy: true);
                    if (tempCT != null)
                    {
                        // Get the name of the existing CT
                        var name = tempCT.EnsureProperty(ct => ct.Name);

                        // If the CT does not exist in the target list, and we don't have to remove it
                        if (!createdList.ContentTypeExistsByName(name) && !ctBinding.Remove)
                        {
                            // Then add it to the target list
                            createdList.AddContentTypeToListById(ctBinding.ContentTypeId, searchContentTypeInSiteHierarchy: true);
                        }
                        // Else if the CT exists in the target list, and we have to remove it
                        else if (createdList.ContentTypeExistsByName(name) && ctBinding.Remove)
                        {
                            // Then remove it from the target list
                            createdList.RemoveContentTypeByName(name);
                        }

                        if (ctBinding.Default)
                        {
                            defaultCtBinding = ctBinding;
                        }
                    }
                }

                // default ContentTypeBinding should be set last because 
                // list extension .SetDefaultContentTypeToList() re-sets 
                // the list.RootFolder UniqueContentTypeOrder property
                // which may cause missing CTs from the "New Button"
                if (defaultCtBinding != null)
                {
                    createdList.SetDefaultContentTypeToList(defaultCtBinding.ContentTypeId);
                }

                // Effectively remove existing content types, if any
                foreach (var ct in contentTypesToRemove)
                {
                    var shouldDelete = true;
                    shouldDelete &= ((createdList.BaseTemplate != (int)ListTemplateType.DocumentLibrary
                        && createdList.BaseTemplate != 851)
                        || !ct.StringId.StartsWith(BuiltInContentTypeId.Folder + "00"));

                    if (shouldDelete)
                    {
                        ct.DeleteObject();
                        web.Context.ExecuteQueryRetry();
                    }
                }
            }

            // Add any custom action
            if (list.UserCustomActions.Any())
            {
                if (!isNoScriptSite)
                {
                    foreach (var userCustomAction in list.UserCustomActions)
                    {
                        CreateListCustomAction(createdList, parser, userCustomAction);
                    }

                    web.Context.ExecuteQueryRetry();
                }
                else
                {
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_SkipAddingOrUpdatingCustomActions);
                }
            }

            if (list.Security != null)
            {
                createdList.SetSecurity(parser, list.Security);
            }
            return Tuple.Create(createdList, parser);
        }

        private void CreateFolderInList(Microsoft.SharePoint.Client.Folder parentFolder, Model.Folder folder, TokenParser parser, PnPMonitoredScope scope)
        {
            // Determine the folder name, parsing any token
            String targetFolderName = parser.ParseString(folder.Name);

            // Check if the folder already exists
            if (parentFolder.FolderExists(targetFolderName))
            {
                // Log a warning if the folder already exists
                String warningFolderAlreadyExists = String.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_FolderAlreadyExists, targetFolderName, parentFolder.ServerRelativeUrl);
                scope.LogWarning(warningFolderAlreadyExists);
                WriteMessage(warningFolderAlreadyExists, ProvisioningMessageType.Warning);
            }

            // Create it or get a reference to it
            var currentFolder = parentFolder.EnsureFolder(targetFolderName);

            if (currentFolder != null)
            {
                // Handle any child-folder
                if (folder.Folders != null && folder.Folders.Count > 0)
                {
                    foreach (var childFolder in folder.Folders)
                    {
                        CreateFolderInList(currentFolder, childFolder, parser, scope);
                    }
                }

                // Handle current folder security
                if (folder.Security != null && folder.Security.RoleAssignments.Count != 0)
                {
                    var currentFolderItem = currentFolder.ListItemAllFields;
                    parentFolder.Context.Load(currentFolderItem);
                    parentFolder.Context.ExecuteQueryRetry();
                    currentFolderItem.SetSecurity(parser, folder.Security);
                }
            }
        }

        private class ListInfo
        {
            public List SiteList { get; set; }
            public ListInstance TemplateList { get; set; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);

                var serverRelativeUrl = web.ServerRelativeUrl;


                // For each list in the site
                var lists = web.Lists;

                web.Context.Load(lists,
                    lc => lc.IncludeWithDefaultProperties(
                        l => l.ContentTypes,
                        l => l.Views,
                        l => l.EnableModeration,
                        l => l.ForceCheckout,
                        l => l.BaseTemplate,
                        l => l.OnQuickLaunch,
                        l => l.RootFolder.ServerRelativeUrl,
                        l => l.UserCustomActions,
                        l => l.MajorVersionLimit,
                        l => l.MajorWithMinorVersionsLimit,
                        l => l.DraftVersionVisibility,
                        l => l.DocumentTemplateUrl,
                        l => l.Fields.IncludeWithDefaultProperties(
                            f => f.Id,
                            f => f.Title,
                            f => f.Hidden,
                            f => f.InternalName,
                            f => f.DefaultValue,
                            f => f.Required)));

                web.Context.ExecuteQueryRetry();

                var allLists = new List<List>();

                if (web.IsSubSite())
                {
                    // If current web is subweb then include the lists in the rootweb for lookup column support
                    var rootWeb = (web.Context as ClientContext).Site.RootWeb;
                    rootWeb.Context.Load(rootWeb.Lists, lsts => lsts.Include(l => l.Id, l => l.Title));
                    rootWeb.Context.ExecuteQueryRetry();
                    foreach (var rootList in rootWeb.Lists)
                    {
                        allLists.Add(rootList);
                    }
                }

                foreach (var list in lists)
                {
                    allLists.Add(list);
                }
                // Let's see if there are workflow subscriptions
                Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription[] workflowSubscriptions = null;
                try
                {
                    workflowSubscriptions = web.GetWorkflowSubscriptions();
                }
                catch (ServerException)
                {
                    // If there is no workflow service present in the farm this method will throw an error. 
                    // Swallow the exception
                }

                // Retrieve all not hidden lists and the Workflow History Lists, just in case there are active workflow subscriptions
                var listsToProcess = lists.AsEnumerable().Where(l => (l.Hidden == false || ((workflowSubscriptions != null && workflowSubscriptions.Length > 0) && l.BaseTemplate == 140))).ToArray();
                var listCount = 0;
                foreach (var siteList in listsToProcess)
                {
                    listCount++;
                    WriteMessage($"List|{siteList.Title}|{listCount}|{listsToProcess.Count()}", ProvisioningMessageType.Progress);
                    ListInstance baseTemplateList = null;
                    if (creationInfo.BaseTemplate != null)
                    {
                        // Check if we need to skip this list...if so let's do it before we gather all the other information for this list...improves performance
                        var index = creationInfo.BaseTemplate.Lists.FindIndex(f => f.Url.Equals(siteList.RootFolder.ServerRelativeUrl.Substring(serverRelativeUrl.Length + 1)) &&
                                                                                   f.TemplateType.Equals(siteList.BaseTemplate));
                        if (index != -1)
                        {
                            baseTemplateList = creationInfo.BaseTemplate.Lists[index];
                        }
                    }

                    var contentTypeFields = new List<FieldRef>();
                    var list = new ListInstance
                    {
                        Description = siteList.Description,
                        EnableVersioning = siteList.EnableVersioning,
                        TemplateType = siteList.BaseTemplate,
                        Title = siteList.Title,
                        Hidden = siteList.Hidden,
                        EnableFolderCreation = siteList.EnableFolderCreation,
                        DocumentTemplate = Tokenize(siteList.DocumentTemplateUrl, web.Url),
                        ContentTypesEnabled = siteList.ContentTypesEnabled,
                        Url = siteList.RootFolder.ServerRelativeUrl.Substring(serverRelativeUrl.Length).TrimStart('/'),
                        TemplateFeatureID = siteList.TemplateFeatureId,
                        EnableAttachments = siteList.EnableAttachments,
                        OnQuickLaunch = siteList.OnQuickLaunch,
                        EnableModeration = siteList.EnableModeration,
                        MaxVersionLimit =
                            siteList.IsPropertyAvailable("MajorVersionLimit") ? siteList.MajorVersionLimit : 0,
                        EnableMinorVersions = siteList.EnableMinorVersions,
                        MinorVersionLimit =
                            siteList.IsPropertyAvailable("MajorWithMinorVersionsLimit")
                                ? siteList.MajorWithMinorVersionsLimit
                                : 0,
                        ForceCheckout = siteList.IsPropertyAvailable("ForceCheckout") ?
                            siteList.ForceCheckout : false,
                        DraftVersionVisibility = siteList.IsPropertyAvailable("DraftVersionVisibility") ? (int)siteList.DraftVersionVisibility : 0,
                    };

                    if (creationInfo.PersistMultiLanguageResources)
                    {
#if !SP2013
                        var escapedListTitle = siteList.Title.Replace(" ", "_");
                        if (UserResourceExtensions.PersistResourceValue(siteList.TitleResource, $"List_{escapedListTitle}_Title", template, creationInfo))
                        {
                            list.Title = $"{{res:List_{escapedListTitle}_Title}}";
                        }
                        if (UserResourceExtensions.PersistResourceValue(siteList.DescriptionResource, $"List_{escapedListTitle}_Description", template, creationInfo))
                        {
                            list.Description = $"{{res:List_{escapedListTitle}_Description}}";
                        }
#endif
                    }

                    list = ExtractContentTypes(web, siteList, contentTypeFields, list);

                    list = ExtractViews(siteList, list);

                    list = ExtractFields(web, siteList, contentTypeFields, list, allLists, creationInfo, template);

                    list = ExtractUserCustomActions(web, siteList, list, creationInfo, template);

                    list.Security = siteList.GetSecurity();

                    if (baseTemplateList != null)
                    {
                        if (!baseTemplateList.Equals(list))
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Adding_list___0_____1_, list.Title, list.Url);
                            template.Lists.Add(list);
                        }
                    }
                    else
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Adding_list___0_____1_, list.Title, list.Url);
                        template.Lists.Add(list);
                    }
                }

            }
            WriteMessage("Done processing lists", ProvisioningMessageType.Completed);
            return template;
        }

        private static ListInstance ExtractViews(List siteList, ListInstance list)
        {
            foreach (var view in siteList.Views.AsEnumerable().Where(view => !view.Hidden))
            {
                var schemaElement = XElement.Parse(view.ListViewXml);

                // Toolbar is not supported

                var toolbarElement = schemaElement.Descendants("Toolbar").FirstOrDefault();
                if (toolbarElement != null)
                {
                    toolbarElement.Remove();
                }

                // XslLink is not supported
                var xslLinkElement = schemaElement.Descendants("XslLink").FirstOrDefault();
                if (xslLinkElement != null)
                {
                    xslLinkElement.Remove();
                }

                list.Views.Add(new View { SchemaXml = schemaElement.ToString() });
            }

            return list;
        }

        private static ListInstance ExtractContentTypes(Web web, List siteList, List<FieldRef> contentTypeFields, ListInstance list)
        {
            var count = 0;

            foreach (var ct in siteList.ContentTypes)
            {
                web.Context.Load(ct, c => c.Parent);
                web.Context.ExecuteQueryRetry();

                if (ct.Parent != null)
                {
                    // Removed this - so that we are getting full list of content types and if it's oob content type,
                    // We are taking parent - VesaJ.
                    //if (!BuiltInContentTypeId.Contains(ct.Parent.StringId)) 
                    //{

                    // Exclude System Content Type to prevent getting exception during import
                    if (!ct.Parent.StringId.Equals(BuiltInContentTypeId.System))
                    {
                        list.ContentTypeBindings.Add(new ContentTypeBinding { ContentTypeId = ct.Parent.StringId, Default = count == 0 });
                    }

                    //}
                }
                else
                {
                    list.ContentTypeBindings.Add(new ContentTypeBinding { ContentTypeId = ct.StringId, Default = count == 0 });
                }

                web.Context.Load(ct.FieldLinks);
                web.Context.ExecuteQueryRetry();
                foreach (var fieldLink in ct.FieldLinks)
                {
                    if (!fieldLink.Hidden)
                    {
                        contentTypeFields.Add(new FieldRef() { Id = fieldLink.Id });
                    }
                }
                count++;
            }

            return list;
        }

        private List<string> SpecialFields => new List<string>() { "LikedBy" };

        private ListInstance ExtractFields(Web web, List siteList, List<FieldRef> contentTypeFields, ListInstance list, List<List> lists, ProvisioningTemplateCreationInformation creationInfo, ProvisioningTemplate template)
        {
            Microsoft.SharePoint.Client.FieldCollection siteColumns = null;
            if (web.IsSubSite())
            {
                var siteContext = web.Context.GetSiteCollectionContext();
                var rootWeb = siteContext.Site.RootWeb;
                siteColumns = rootWeb.Fields;
                siteContext.Load(siteColumns, scs => scs.Include(sc => sc.Id, sc => sc.DefaultValue));
                siteContext.ExecuteQueryRetry();
            }
            else
            {
                siteColumns = web.Fields;
                web.Context.Load(siteColumns, scs => scs.Include(sc => sc.Id, sc => sc.DefaultValue));
                web.Context.ExecuteQueryRetry();
            }

            var fieldsToProcess = siteList.Fields.AsEnumerable().Where(field => !field.Hidden || SpecialFields.Contains(field.InternalName)).ToArray();

            foreach (var field in fieldsToProcess)
            {
                var siteColumn = siteColumns.FirstOrDefault(sc => sc.Id == field.Id);
                if (siteColumn != null)
                {
                    var addField = true;
                    if (siteList.ContentTypesEnabled && contentTypeFields.FirstOrDefault(c => c.Id == field.Id) == null)
                    {
                        if (contentTypeFields.FirstOrDefault(c => c.Id == field.Id) == null)
                        {
                            addField = false;
                        }
                    }

                    if (siteColumn.DefaultValue != field.DefaultValue)
                    {
                        list.FieldDefaults.Add(field.InternalName, field.DefaultValue);
                        addField = true;
                    }

                    var fieldElement = XElement.Parse(field.SchemaXml);
                    var sourceId = fieldElement.Attribute("SourceID") != null ? fieldElement.Attribute("SourceID").Value : null;

                    if (sourceId != null && sourceId == "http://schemas.microsoft.com/sharepoint/v3")
                    {
                        if (field.InternalName == "Editor" ||
                            field.InternalName == "Author" ||
                            field.InternalName == "Title" ||
                            field.InternalName == "ID" ||
                            field.InternalName == "Created" ||
                            field.InternalName == "Modified" ||
                            field.InternalName == "Attachments" ||
                            field.InternalName == "_UIVersionString" ||
                            field.InternalName == "DocIcon" ||
                            field.InternalName == "LinkTitleNoMenu" ||
                            field.InternalName == "LinkTitle" ||
                            field.InternalName == "Edit" ||
                            field.InternalName == "AppAuthor" ||
                            field.InternalName == "AppEditor" ||
                            field.InternalName == "ContentType" ||
                            field.InternalName == "ItemChildCount" ||
                            field.InternalName == "FolderChildCount" ||
                            field.InternalName == "LinkFilenameNoMenu" ||
                            field.InternalName == "LinkFilename" ||
                            field.InternalName == "_CopySource" ||
                            field.InternalName == "ParentVersionString" ||
                            field.InternalName == "ParentLeafName" ||
                            field.InternalName == "_CheckinComment" ||
                            field.InternalName == "FileLeafRef" ||
                            field.InternalName == "FileSizeDisplay" ||
                            field.InternalName == "Preview" ||
                            field.InternalName == "ThumbnailOnForm" ||
                            field.InternalName == "CheckoutUser" ||
                            field.InternalName == "Modified_x0020_By" ||
                            field.InternalName == "Created_x0020_By"
                            )
                        {
                            addField = false;
                        }
                    }

                    if (addField)
                    {
                        list.FieldRefs.Add(new FieldRef(field.InternalName)
                        {
                            Id = field.Id,
                            DisplayName = field.Title,
                            Required = field.Required,
                            Hidden = field.Hidden,
                        });
                        if (field.TypeAsString.StartsWith("TaxonomyField"))
                        {
                            // find the corresponding taxonomy field and include it anyway
                            var taxField = (TaxonomyField)field;
                            taxField.EnsureProperties(f => f.TextField, f => f.Id);

                            var noteField = siteList.Fields.GetById(taxField.TextField);
                            web.Context.Load(noteField, nf => nf.Id, nf => nf.Title, nf => nf.Required, nf => nf.Hidden, nf => nf.InternalName);
                            web.Context.ExecuteQueryRetry();

                            list.FieldRefs.Insert(0, new FieldRef(noteField.InternalName)
                            {
                                Id = noteField.Id,
                                DisplayName = noteField.Title,
                                Required = noteField.Required,
                                Hidden = noteField.Hidden
                            });
                        }
                    }
                }
                else
                {
                    var schemaXml = ParseFieldSchema(field.SchemaXml, web, lists);
                    var fieldElement = XElement.Parse(field.SchemaXml);
                    var listId = fieldElement.Attribute("List") != null ? fieldElement.Attribute("List").Value : null;

                    if (fieldElement.Attribute("Type").Value == "Calculated")
                    {
                        schemaXml = ObjectField.TokenizeFieldFormula(siteList.Fields, (FieldCalculated)field, schemaXml);
                    }

                    //Field has column Validation
                    if (fieldElement.Elements("Validation").FirstOrDefault() != null)
                    {
                        schemaXml = ObjectField.TokenizeFieldValidationFormula(field, schemaXml);
                    }

                    if (creationInfo.PersistMultiLanguageResources)
                    {
#if !SP2013
                        var escapedFieldTitle = field.Title.Replace(" ", "_");
                        if (UserResourceExtensions.PersistResourceValue(field.TitleResource, $"Field_{escapedFieldTitle}_DisplayName", template, creationInfo))
                        {
                            var fieldTitle = $"{{res:Field_{escapedFieldTitle}_DisplayName}}";
                            fieldElement.SetAttributeValue("DisplayName", fieldTitle);

                        }
                        if (UserResourceExtensions.PersistResourceValue(field.DescriptionResource, $"Field_{escapedFieldTitle}_Description", template, creationInfo))
                        {
                            var fieldDescription = $"{{res:Field_{escapedFieldTitle}_Description}}";
                            fieldElement.SetAttributeValue("Description", fieldDescription);
                        }

                        schemaXml = fieldElement.ToString();
#endif
                    }

                    if (listId == null)
                    {
                        list.Fields.Add((new Model.Field { SchemaXml = schemaXml }));
                    }
                    else
                    {
                        var listIdValue = Guid.Empty;
                        if (Guid.TryParse(listId, out listIdValue))
                        {
                            var sourceList = lists.AsEnumerable().Where(l => l.Id == listIdValue).FirstOrDefault();
                            if (sourceList != null)
                                fieldElement.Attribute("List").SetValue($"{{listid:{sourceList.Title}}}");
                        }
                        var fieldSchema = fieldElement.ToString();
                        if (field.TypeAsString.StartsWith("TaxonomyField"))
                        {
                            fieldSchema = TokenizeTaxonomyField(web, fieldElement);
                        }
                        list.Fields.Add(new Model.Field { SchemaXml = ParseFieldSchema(fieldSchema, web, lists) });
                    }

                    if (field.TypeAsString.StartsWith("TaxonomyField"))
                    {
                        // find the corresponding taxonomy container text field and include it too
                        var taxField = (TaxonomyField)field;
                        taxField.EnsureProperties(f => f.TextField, f => f.Id);

                        var noteField = siteList.Fields.GetById(taxField.TextField);
                        web.Context.Load(noteField, nf => nf.SchemaXml);
                        web.Context.ExecuteQueryRetry();
                        var noteSchemaXml = XElement.Parse(noteField.SchemaXml);
                        noteSchemaXml.Attribute("SourceID")?.Remove();
                        list.Fields.Insert(0, new Model.Field { SchemaXml = ParseFieldSchema(noteSchemaXml.ToString(), web, lists) });
                    }

                }
            }
            return list;
        }

        private static ListInstance ExtractUserCustomActions(Web web, List siteList, ListInstance list, ProvisioningTemplateCreationInformation creationInfo, ProvisioningTemplate template)
        {
            foreach (var userCustomAction in siteList.UserCustomActions.AsEnumerable())
            {
                web.Context.Load(userCustomAction);
                web.Context.ExecuteQueryRetry();

                var customAction = new CustomAction
                {
                    Title = userCustomAction.Title,
                    Description = userCustomAction.Description,
                    Enabled = true,
                    Name = userCustomAction.Name,
                    //RegistrationType = userCustomAction.RegistrationType,
                    //RegistrationId = userCustomAction.RegistrationId,
                    Url = userCustomAction.Url,
                    ImageUrl = userCustomAction.ImageUrl,
                    Rights = userCustomAction.Rights,
                    Sequence = userCustomAction.Sequence,
                    ScriptBlock = userCustomAction.ScriptBlock,
                    ScriptSrc = userCustomAction.ScriptSrc,
                    CommandUIExtension = !System.String.IsNullOrEmpty(userCustomAction.CommandUIExtension) ?
                        XElement.Parse(userCustomAction.CommandUIExtension) : null,
                    Group = userCustomAction.Group,
                    Location = userCustomAction.Location,
                };

#if !ONPREMISES
                if (creationInfo.PersistMultiLanguageResources)
                {
                    siteList.EnsureProperty(l => l.Title);
                    var listKey = siteList.Title.Replace(" ", "_");
                    var resourceKey = userCustomAction.Name.Replace(" ", "_");

                    if (UserResourceExtensions.PersistResourceValue(userCustomAction.TitleResource, $"List_{listKey}_CustomAction_{resourceKey}_Title", template, creationInfo))
                    {
                        var customActionTitle = $"{{res:List_{listKey}_CustomAction_{resourceKey}_Title}}";
                        customAction.Title = customActionTitle;

                    }
                    if (UserResourceExtensions.PersistResourceValue(userCustomAction.DescriptionResource, $"List_{listKey}_CustomAction_{resourceKey}_Description", template, creationInfo))
                    {
                        var customActionDescription = $"{{res:List_{listKey}_CustomAction_{resourceKey}_Description}}";
                        customAction.Description = customActionDescription;
                    }
                }
#endif

                list.UserCustomActions.Add(customAction);
            }

            return list;
        }

        private string ParseFieldSchema(string schemaXml, Web web, List<List> lists)
        {
            foreach (var list in lists)
            {
                schemaXml = Regex.Replace(schemaXml, list.Id.ToString(), $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}", RegexOptions.IgnoreCase);
            }
            schemaXml = Regex.Replace(schemaXml, web.Id.ToString("B"), "{{siteid}}", RegexOptions.IgnoreCase);
            schemaXml = Regex.Replace(schemaXml, web.Id.ToString("D"), "{siteid}", RegexOptions.IgnoreCase);
            return schemaXml;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Lists.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                var collList = web.Lists;
                var lists = web.Context.LoadQuery(collList.Where(l => l.Hidden == false));

                web.Context.ExecuteQueryRetry();

                _willExtract = lists.Any();
            }
            return _willExtract.Value;
        }
    }
}
