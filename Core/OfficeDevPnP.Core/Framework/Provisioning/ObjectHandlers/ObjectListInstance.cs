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
                    var rootWeb = (web.Context as ClientContext).Site.RootWeb;

                    web.EnsureProperties(w => w.ServerRelativeUrl);

                    web.Context.Load(web.Lists, lc => lc.IncludeWithDefaultProperties(l => l.RootFolder.ServerRelativeUrl));
                    web.Context.ExecuteQueryRetry();
                    var existingLists = web.Lists.AsEnumerable().Select(existingList => existingList.RootFolder.ServerRelativeUrl).ToList();
                    var serverRelativeUrl = web.ServerRelativeUrl;

                    var processedLists = new List<ListInfo>();

                    #region Lists

                    foreach (var templateList in template.Lists)
                    {
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
                                    throw new Exception(string.Format("Referenced content type {0} not available in site or in template", ct.ContentTypeId));
                                }
                            }
                        }
                        var index = existingLists.FindIndex(x => x.Equals(UrlUtility.Combine(serverRelativeUrl, templateList.Url), StringComparison.OrdinalIgnoreCase));
                        if (index == -1)
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_list__0_, templateList.Title);
                                var returnTuple = CreateList(web, templateList, parser, scope);
                                var createdList = returnTuple.Item1;
                                parser = returnTuple.Item2;
                                processedLists.Add(new ListInfo { SiteList = createdList, TemplateList = templateList });

                                parser.AddToken(new ListIdToken(web, templateList.Title, createdList.Id));

                                parser.AddToken(new ListUrlToken(web, templateList.Title, createdList.RootFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length + 1)));
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
                                var returnTuple = UpdateList(web, existingList, templateList, parser, scope);
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

                            foreach (var fieldRef in listInfo.TemplateList.FieldRefs)
                            {
                                var field = rootWeb.GetFieldById<Field>(fieldRef.Id);
                                if (field != null)
                                {
                                    if (!listInfo.SiteList.FieldExistsById(fieldRef.Id))
                                    {
                                        CreateFieldRef(listInfo, field, fieldRef);
                                    }
                                    else
                                    {
                                        UpdateFieldRef(listInfo.SiteList, field.Id, fieldRef);
                                    }
                                }

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
                            foreach (var field in listInfo.TemplateList.Fields)
                            {
                                var fieldElement = XElement.Parse(parser.ParseString(field.SchemaXml, "~sitecollection", "~site"));
                                if (fieldElement.Attribute("ID") == null)
                                {
                                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field_schema_has_no_ID_attribute___0_, field.SchemaXml);
                                    throw new Exception(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field_schema_has_no_ID_attribute___0_, field.SchemaXml));
                                }
                                var id = fieldElement.Attribute("ID").Value;

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
                                            CreateField(fieldElement, listInfo, parser);
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
                                            UpdateField(web, listInfo, fieldGuid, fieldElement, fieldFromList, scope, parser);
                                        }
                                        catch (Exception ex)
                                        {
                                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Updating_field__0__failed___1_____2_, fieldGuid, ex.Message, ex.StackTrace);
                                            throw;
                                        }

                                    }
                                }
                            }
                        }
                        listInfo.SiteList.Update();
                        web.Context.ExecuteQueryRetry();
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
                        foreach (var view in list.Views)
                        {

                            CreateView(web, view, existingViews, createdList, scope);

                        }

                        //// Removing existing views set the OnQuickLaunch option to false and need to be re-set.
                        //if (list.OnQuickLaunch && list.RemoveExistingViews && list.Views.Count > 0)
                        //{
                        //    createdList.RefreshLoad();
                        //    web.Context.ExecuteQueryRetry();
                        //    createdList.OnQuickLaunch = list.OnQuickLaunch;
                        //    createdList.Update();
                        //    web.Context.ExecuteQueryRetry();
                        //}
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

                }
            }
            return parser;
        }

        private void CreateView(Web web, View view, Microsoft.SharePoint.Client.ViewCollection existingViews, List createdList, PnPMonitoredScope monitoredScope)
        {
            try
            {

                var viewElement = XElement.Parse(view.SchemaXml);
                var displayNameElement = viewElement.Attribute("DisplayName");
                if (displayNameElement == null)
                {
                    throw new ApplicationException("Invalid View element, missing a valid value for the attribute DisplayName.");
                }

                monitoredScope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_view__0_, displayNameElement.Value);
                var existingView = existingViews.FirstOrDefault(v => v.Title == displayNameElement.Value);

                if (existingView != null)
                {
                    existingView.DeleteObject();
                    web.Context.ExecuteQueryRetry();
                }

                var viewTitle = displayNameElement.Value;

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
                var viewDefault = viewElement.Attribute("DefaultView") != null && Boolean.Parse(viewElement.Attribute("DefaultView").Value);

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
                web.Context.Load(createdView, v => v.Scope, v => v.JSLink, v => v.Title);
                web.Context.ExecuteQueryRetry();

                if (urlHasValue)
                {
                    //restore original title 
                    createdView.Title = viewTitle;
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

                // JSLink
                var jslinkElement = viewElement.Descendants("JSLink").FirstOrDefault();
                if (jslinkElement != null)
                {
                    var jslink = jslinkElement.Value;
                    if (createdView.JSLink != jslink)
                    {
                        createdView.JSLink = jslink;
                        createdView.Update();
                    }
                }

                createdList.Update();
                web.Context.ExecuteQueryRetry();
            }
            catch (Exception ex)
            {
                monitoredScope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Creating_view_failed___0_____1_, ex.Message, ex.StackTrace);
                throw;
            }
        }

        private static void UpdateFieldRef(List siteList, Guid fieldId, FieldRef fieldRef)
        {
            // find the field in the list
            var listField = siteList.Fields.GetById(fieldId);

            siteList.Context.Load(listField, f => f.Title, f => f.Hidden, f => f.Required);
            siteList.Context.ExecuteQueryRetry();

            var isDirty = false;
            if (!string.IsNullOrEmpty(fieldRef.DisplayName) && fieldRef.DisplayName != listField.Title)
            {
                listField.Title = fieldRef.DisplayName;
                isDirty = true;
            }
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
        }

        private static void CreateFieldRef(ListInfo listInfo, Field field, FieldRef fieldRef)
        {
            XElement element = XElement.Parse(field.SchemaXml);

            element.SetAttributeValue("AllowDeletion", "TRUE");

            field.SchemaXml = element.ToString();

            var createdField = listInfo.SiteList.Fields.Add(field);

            createdField.Context.Load(createdField, cf => cf.Title, cf => cf.Hidden, cf => cf.Required);
            createdField.Context.ExecuteQueryRetry();

            var isDirty = false;
            if (!string.IsNullOrEmpty(fieldRef.DisplayName) && createdField.Title != fieldRef.DisplayName)
            {
                createdField.Title = fieldRef.DisplayName;
                isDirty = true;
            }
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
        }

        private static void CreateField(XElement fieldElement, ListInfo listInfo, TokenParser parser)
        {
            fieldElement = PrepareField(fieldElement);

            var fieldXml = parser.ParseString(fieldElement.ToString(), "~sitecollection", "~site");
            listInfo.SiteList.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddFieldInternalNameHint);
            listInfo.SiteList.Context.ExecuteQueryRetry();
        }

        private void UpdateField(ClientObject web, ListInfo listInfo, Guid fieldId, XElement templateFieldElement, Field existingField, PnPMonitoredScope scope, TokenParser parser)
        {
            web.Context.Load(existingField, f => f.SchemaXml);
            web.Context.ExecuteQueryRetry();

            var existingFieldElement = XElement.Parse(existingField.SchemaXml);

            var equalityComparer = new XNodeEqualityComparer();

            // Is field different in template?
            if (equalityComparer.GetHashCode(existingFieldElement) != equalityComparer.GetHashCode(templateFieldElement))
            {
                // Is existing field of the same type?
                if (existingFieldElement.Attribute("Type").Value == templateFieldElement.Attribute("Type").Value)
                {
                    templateFieldElement = PrepareField(templateFieldElement);

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

                    if (existingFieldElement.Attribute("Version") != null)
                    {
                        existingFieldElement.Attributes("Version").Remove();
                    }
                    existingField.SchemaXml = parser.ParseString(existingFieldElement.ToString(), "~sitecollection", "~site");
                    existingField.UpdateAndPushChanges(true);
                    web.Context.ExecuteQueryRetry();
                }
                else
                {
                    var fieldName = existingFieldElement.Attribute("Name") != null ? existingFieldElement.Attribute("Name").Value : existingFieldElement.Attribute("StaticName").Value;
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field__0____1___exists_in_list__2____3___but_is_of_different_type__Skipping_field_, fieldName, fieldId, listInfo.TemplateList.Title, listInfo.SiteList.Id);
                    WriteWarning(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field__0____1___exists_in_list__2____3___but_is_of_different_type__Skipping_field_, fieldName, fieldId, listInfo.TemplateList.Title, listInfo.SiteList.Id), ProvisioningMessageType.Warning);
                }
            }
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

        private Tuple<List, TokenParser> UpdateList(Web web, List existingList, ListInstance templateList, TokenParser parser, PnPMonitoredScope scope)
        {
            web.Context.Load(existingList,
                l => l.Title,
                l => l.Description,
                l => l.OnQuickLaunch,
                l => l.Hidden,
                l => l.ContentTypesEnabled,
                l => l.EnableAttachments,
                l => l.EnableFolderCreation,
                l => l.EnableMinorVersions,
                l => l.DraftVersionVisibility,
                l => l.Views,
                l => l.RootFolder
#if !CLIENTSDKV15
, l => l.MajorWithMinorVersionsLimit
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
                if (existingList.BaseTemplate != (int)ListTemplateType.Survey && existingList.BaseTemplate != (int)ListTemplateType.DocumentLibrary)
                {
                    // https://msdn.microsoft.com/EN-US/library/microsoft.sharepoint.splist.enableattachments.aspx
                    // The EnableAttachments property does not apply to any list that has a base type of Survey or DocumentLibrary.
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
                if (templateList.EnableVersioning)
                {
                    if (existingList.EnableVersioning != templateList.EnableVersioning)
                    {
                        existingList.EnableVersioning = templateList.EnableVersioning;
                        isDirty = true;
                    }
#if !CLIENTSDKV15
                    if (existingList.IsObjectPropertyInstantiated("MajorVersionLimit") && existingList.MajorVersionLimit != templateList.MaxVersionLimit)
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
                if (isDirty)
                {
                    existingList.Update();
                    web.Context.ExecuteQueryRetry();
                }


                if (existingList.ContentTypesEnabled)
                {
                    // Check if we need to add a content type

                    var existingContentTypes = existingList.ContentTypes;
                    web.Context.Load(existingContentTypes, cts => cts.Include(ct => ct.StringId));
                    web.Context.ExecuteQueryRetry();

                    var bindingsToAdd = templateList.ContentTypeBindings.Where(ctb => existingContentTypes.All(ct => !ctb.ContentTypeId.Equals(ct.StringId, StringComparison.InvariantCultureIgnoreCase))).ToList();
                    var defaultCtBinding = templateList.ContentTypeBindings.FirstOrDefault(ctb => ctb.Default == true);

                    var bindingAddedToList = false;
                    foreach (var ctb in bindingsToAdd)
                    {
                        // Added a check so that if no bindings were actually added then the SetDefaultContentTypeToList method will not be executed
                        // This is to address a specific scenario when OOTB PWA lists can not be updated as they are centrally managed
                        var addedToList = existingList.AddContentTypeToListById(ctb.ContentTypeId, searchContentTypeInSiteHierarchy: true);
                        if (addedToList)
                        {
                            bindingAddedToList = true;
                        }
                    }

                    // default ContentTypeBinding should be set last because 
                    // list extension .SetDefaultContentTypeToList() re-sets 
                    // the list.RootFolder UniqueContentTypeOrder property
                    // which may cause missing CTs from the "New Button"
                    if (defaultCtBinding != null && bindingAddedToList)
                    {
                        existingList.SetDefaultContentTypeToList(defaultCtBinding.ContentTypeId);
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
                WriteWarning(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_List__0____1____2___exists_but_is_of_a_different_type__Skipping_list_, templateList.Title, templateList.Url, existingList.Id), ProvisioningMessageType.Warning);
                return null;
            }
        }

        private Tuple<List, TokenParser> CreateList(Web web, ListInstance list, TokenParser parser, PnPMonitoredScope scope)
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

            var createdList = web.Lists.Add(listCreate);
            createdList.Update();
            web.Context.Load(createdList, l => l.BaseTemplate);
            web.Context.ExecuteQueryRetry();

            if (!String.IsNullOrEmpty(list.DocumentTemplate))
            {
                createdList.DocumentTemplateUrl = parser.ParseString(list.DocumentTemplate);
            }

            // EnableAttachments are not supported for DocumentLibraries and Surveys
            // TODO: the user should be warned
            if (createdList.BaseTemplate != (int)ListTemplateType.DocumentLibrary && createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                createdList.EnableAttachments = list.EnableAttachments;
            }

            createdList.EnableModeration = list.EnableModeration;

            // Done for all other lists than for Survey - With Surveys versioning configuration will cause an exception
            if (createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                createdList.EnableVersioning = list.EnableVersioning;
                if (list.EnableVersioning)
                {
#if !CLIENTSDKV15
                    createdList.MajorVersionLimit = list.MaxVersionLimit;
#endif

                    if (createdList.BaseTemplate == (int)ListTemplateType.DocumentLibrary)
                    {
                        // Only supported on Document Libraries
                        createdList.EnableMinorVersions = list.EnableMinorVersions;
                        createdList.DraftVersionVisibility = (DraftVisibilityType)list.DraftVersionVisibility;

                        if (list.EnableMinorVersions)
                        {
                            createdList.MajorWithMinorVersionsLimit = list.MinorVersionLimit; // Set only if enabled, otherwise you'll get exception due setting value to zero.

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
                                    WriteWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_DraftVersionVisibility_not_applied_because_EnableModeration_is_not_set_to_true, ProvisioningMessageType.Warning);
                                }
                            }
                            else
                            {
                                createdList.DraftVersionVisibility = (DraftVisibilityType)list.DraftVersionVisibility;
                            }
                        }
                    }
                }
            }

            createdList.OnQuickLaunch = list.OnQuickLaunch;
            if (createdList.BaseTemplate != (int)ListTemplateType.DiscussionBoard)
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
                        // Check if CT is already available
                        var name = tempCT.EnsureProperty(ct => ct.Name);
                        if (!createdList.ContentTypeExistsByName(name))
                        {
                            createdList.AddContentTypeToListById(ctBinding.ContentTypeId, searchContentTypeInSiteHierarchy: true);
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
                    shouldDelete &= (createdList.BaseTemplate != (int)ListTemplateType.DocumentLibrary || !ct.StringId.StartsWith(BuiltInContentTypeId.Folder + "00"));

                    if (shouldDelete)
                    {
                        ct.DeleteObject();
                        web.Context.ExecuteQueryRetry();
                    }
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
                WriteWarning(warningFolderAlreadyExists, ProvisioningMessageType.Warning);
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
                        l => l.BaseTemplate,
                        l => l.OnQuickLaunch,
                        l => l.RootFolder.ServerRelativeUrl,
                        l => l.Fields.IncludeWithDefaultProperties(
                            f => f.Id,
                            f => f.Title,
                            f => f.Hidden,
                            f => f.InternalName,
                            f => f.Required)));

                web.Context.ExecuteQueryRetry();

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
                foreach (var siteList in lists.AsEnumerable().Where(l => (l.Hidden == false || ((workflowSubscriptions != null && workflowSubscriptions.Length > 0) && l.BaseTemplate == 140))))
                {
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
                        MaxVersionLimit =
                            siteList.IsObjectPropertyInstantiated("MajorVersionLimit") ? siteList.MajorVersionLimit : 0,
                        EnableMinorVersions = siteList.EnableMinorVersions,
                        MinorVersionLimit =
                            siteList.IsObjectPropertyInstantiated("MajorWithMinorVersionsLimit")
                                ? siteList.MajorWithMinorVersionsLimit
                                : 0
                    };


                    list = ExtractContentTypes(web, siteList, contentTypeFields, list);

                    list = ExtractViews(siteList, list);

                    list = ExtractFields(web, siteList, contentTypeFields, list, lists);

                    list.Security = siteList.GetSecurity();

                    var logCTWarning = false;
                    if (baseTemplateList != null)
                    {
                        if (!baseTemplateList.Equals(list))
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Adding_list___0_____1_, list.Title, list.Url);
                            template.Lists.Add(list);
                            if (list.ContentTypesEnabled && list.ContentTypeBindings.Any() && web.IsSubSite())
                            {
                                logCTWarning = true;
                            }
                        }
                    }
                    else
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstances_Adding_list___0_____1_, list.Title, list.Url);
                        template.Lists.Add(list);
                        if (list.ContentTypesEnabled && list.ContentTypeBindings.Any() && web.IsSubSite())
                        {
                            logCTWarning = true;
                        }

                    }
                    if (logCTWarning)
                    {
                        scope.LogWarning("You are extracting a template from a subweb. List '{0}' refers to content types. Content types are not exported when extracting a template from a subweb", list.Title);
                        WriteWarning(string.Format("You are extracting a template from a subweb. List '{0}' refers to content types. Content types are not exported when extracting a template from a subweb", list.Title), ProvisioningMessageType.Warning);
                    }
                }

            }
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
                    list.ContentTypeBindings.Add(new ContentTypeBinding { ContentTypeId = ct.Parent.StringId, Default = count == 0 });
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

        private ListInstance ExtractFields(Web web, List siteList, List<FieldRef> contentTypeFields, ListInstance list, ListCollection lists)
        {
            var siteColumns = web.Fields;
            web.Context.Load(siteColumns, scs => scs.Include(sc => sc.Id));
            web.Context.ExecuteQueryRetry();

            foreach (var field in siteList.Fields.AsEnumerable().Where(field => !field.Hidden))
            {
                if (siteColumns.FirstOrDefault(sc => sc.Id == field.Id) != null)
                {
                    var addField = true;
                    if (siteList.ContentTypesEnabled && contentTypeFields.FirstOrDefault(c => c.Id == field.Id) == null)
                    {
                        if (contentTypeFields.FirstOrDefault(c => c.Id == field.Id) == null)
                        {
                            addField = false;
                        }
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
                            field.InternalName == "ThumbnailOnForm")
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
                    var schemaXml = ParseFieldSchema(field.SchemaXml, lists);
                    var fieldElement = XElement.Parse(field.SchemaXml);
                    var listId = fieldElement.Attribute("List") != null ? fieldElement.Attribute("List").Value : null;

                    if (listId == null)
                        list.Fields.Add((new Model.Field { SchemaXml = field.SchemaXml }));
                    else
                    {
                        var listIdValue = Guid.Empty;
                        if (Guid.TryParse(listId, out listIdValue))
                        {
                            var sourceList = lists.AsEnumerable().Where(l => l.Id == listIdValue).FirstOrDefault();
                            if (sourceList != null)
                                fieldElement.Attribute("List").SetValue(String.Format("{{listid:{0}}}", sourceList.Title));
                        }

                        list.Fields.Add(new Model.Field { SchemaXml = fieldElement.ToString() });
                    }

                    if (field.TypeAsString.StartsWith("TaxonomyField"))
                    {
                        // find the corresponding taxonomy field and include it anyway
                        var taxField = (TaxonomyField)field;
                        taxField.EnsureProperties(f => f.TextField, f => f.Id);

                        var noteField = siteList.Fields.GetById(taxField.TextField);
                        web.Context.Load(noteField, nf => nf.SchemaXml);
                        web.Context.ExecuteQueryRetry();
                        var noteSchemaXml = XElement.Parse(noteField.SchemaXml);
                        noteSchemaXml.Attribute("SourceID").Remove();
                        list.Fields.Insert(0, new Model.Field { SchemaXml = ParseFieldSchema(noteSchemaXml.ToString(), lists) });
                    }

                }
            }
            return list;
        }

        private string ParseFieldSchema(string schemaXml, ListCollection lists)
        {
            foreach (var list in lists)
            {
                schemaXml = Regex.Replace(schemaXml, list.Id.ToString(), string.Format("{{listid:{0}}}", list.Title), RegexOptions.IgnoreCase);
            }

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
