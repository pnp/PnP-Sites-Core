using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using Field = Microsoft.SharePoint.Client.Field;
using View = OfficeDevPnP.Core.Framework.Provisioning.Model.View;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectListInstance : ObjectHandlerBase
    {
        private readonly FieldAndListProvisioningStepHelper.Step step;
        public override string Name
        {
#if DEBUG
            get { return $"List instances ({step})"; }
#else
            get { return $"List instances"; }
#endif
        }

        public ObjectListInstance(FieldAndListProvisioningStepHelper.Step stage)
        {
            this.step = stage;
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

                    #region Lists and List Content Types

                    var currentListIndex = 0;
                    foreach (var templateList in template.Lists)
                    {
                        templateList.Url = parser.ParseString(templateList.Url);
                        currentListIndex++;
                        WriteMessage($"List|{templateList.Title}|{currentListIndex}|{total}", ProvisioningMessageType.Progress);
                        CheckContentTypes(web, template, scope, templateList);
                        // check if the List exists by url or by title
                        var index = existingLists.FindIndex(x => x.Title.Equals(parser.ParseString(templateList.Title), StringComparison.OrdinalIgnoreCase) || x.RootFolder.ServerRelativeUrl.Equals(UrlUtility.Combine(serverRelativeUrl, templateList.Url), StringComparison.OrdinalIgnoreCase));

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

                    #endregion Lists and List Content Types

                    #region FieldRefs

                    foreach (var listInfo in processedLists)
                    {
                        ProcessFieldRefs(web, template, parser, scope, rootWeb, listInfo);
                    }

                    #endregion FieldRefs

                    #region Fields

                    foreach (var listInfo in processedLists)
                    {
                        ProcessFields(web, parser, scope, listInfo);
                    }

                    #endregion Fields

                    // We stop here unless we reached the last provisioning stop of the list
                    if (step == FieldAndListProvisioningStepHelper.Step.ListSettings)
                    {

                        #region Default Field Values

                        foreach (var listInfo in processedLists)
                        {
                            ProcessFieldDefaults(web, listInfo);
                        }

                        #endregion Default Field Values

                        #region Views

                        foreach (var listInfo in processedLists)
                        {
                            ProcessViews(web, parser, scope, listInfo);
                        }

                        #endregion Views

                        #region Folders

                        // Folders are supported for document libraries and generic lists only
                        foreach (var list in processedLists)
                        {
                            ProcessFolders(web, parser, scope, list);
                        }

                        #endregion Folders

                        #region IRM Settings

                        // Configure IRM Settings
                        foreach (var list in processedLists)
                        {
                            ProcessIRMSettings(web, list);
                        }

                        #endregion IRM Settings

                        // If an existing view is updated, and the list is to be listed on the QuickLaunch, it is removed because the existing view will be deleted and recreated from scratch.
                        foreach (var listInfo in processedLists)
                        {
                            listInfo.SiteList.OnQuickLaunch = listInfo.TemplateList.OnQuickLaunch;
                            listInfo.SiteList.Update();
                        }
                        web.Context.ExecuteQueryRetry();
                    }
                    WriteMessage("Done processing lists", ProvisioningMessageType.Completed);
                }
            }
            return parser;
        }

        private static void ProcessIRMSettings(Web web, ListInfo list)
        {
            if (list.SiteList.BaseTemplate != (int)ListTemplateType.PictureLibrary && list.TemplateList.IRMSettings != null && list.TemplateList.IRMSettings.Enabled)
            {
                list.SiteList.IrmEnabled = true;
                list.SiteList.IrmExpire = list.TemplateList.IrmExpire;
                list.SiteList.IrmReject = list.TemplateList.IrmReject;

                list.SiteList.InformationRightsManagementSettings.AllowPrint = list.TemplateList.IRMSettings.AllowPrint;
                list.SiteList.InformationRightsManagementSettings.AllowScript = list.TemplateList.IRMSettings.AllowScript;
                list.SiteList.InformationRightsManagementSettings.AllowWriteCopy = list.TemplateList.IRMSettings.AllowWriteCopy;
                list.SiteList.InformationRightsManagementSettings.DisableDocumentBrowserView = list.TemplateList.IRMSettings.DisableDocumentBrowserView;
                list.SiteList.InformationRightsManagementSettings.DocumentAccessExpireDays = list.TemplateList.IRMSettings.DocumentAccessExpireDays;
                if (list.TemplateList.IRMSettings.DocumentLibraryProtectionExpiresInDays > 0)
                {
                    list.SiteList.InformationRightsManagementSettings.DocumentLibraryProtectionExpireDate = DateTime.Now.AddDays(list.TemplateList.IRMSettings.DocumentLibraryProtectionExpiresInDays);
                }
                list.SiteList.InformationRightsManagementSettings.EnableDocumentAccessExpire = list.TemplateList.IRMSettings.EnableDocumentAccessExpire;
                list.SiteList.InformationRightsManagementSettings.EnableDocumentBrowserPublishingView = list.TemplateList.IRMSettings.EnableDocumentBrowserPublishingView;
                list.SiteList.InformationRightsManagementSettings.EnableGroupProtection = list.TemplateList.IRMSettings.EnableGroupProtection;
                list.SiteList.InformationRightsManagementSettings.EnableLicenseCacheExpire = list.TemplateList.IRMSettings.EnableLicenseCacheExpire;
                list.SiteList.InformationRightsManagementSettings.GroupName = list.TemplateList.IRMSettings.GroupName;
                list.SiteList.InformationRightsManagementSettings.LicenseCacheExpireDays = list.TemplateList.IRMSettings.LicenseCacheExpireDays;
                list.SiteList.InformationRightsManagementSettings.PolicyDescription = list.TemplateList.IRMSettings.PolicyDescription;
                list.SiteList.InformationRightsManagementSettings.PolicyTitle = list.TemplateList.IRMSettings.PolicyTitle;

                list.SiteList.Update();
                web.Context.ExecuteQueryRetry();
            }
        }

        private void ProcessFolders(Web web, TokenParser parser, PnPMonitoredScope scope, ListInfo list)
        {
            list.SiteList.EnsureProperties(l => l.BaseType);
            if ((list.SiteList.BaseType == BaseType.DocumentLibrary
                || list.SiteList.BaseType == BaseType.GenericList)
                && list.TemplateList.Folders != null && list.TemplateList.Folders.Count > 0)
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

        private void ProcessViews(Web web, TokenParser parser, PnPMonitoredScope scope, ListInfo listInfo)
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
            var currentViewIndex = 0;
            foreach (var view in list.Views)
            {
                currentViewIndex++;
                CreateView(web, view, existingViews, createdList, scope, parser, currentViewIndex, list.Views.Count);
            }
        }

        private static void ProcessFieldDefaults(Web web, ListInfo listInfo)
        {
            if (listInfo.TemplateList.FieldDefaults.Count > 0)
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

        private void ProcessFields(Web web, TokenParser parser, PnPMonitoredScope scope, ListInfo listInfo)
        {
            if (listInfo.TemplateList.Fields.Any())
            {
                var currentFieldIndex = 0;
                var fieldsToProcess = listInfo.TemplateList.Fields
                    .Select(fld => new
                    {
                        Field = fld,
                        FieldRef = (string)XElement.Parse(parser.ParseString(fld.SchemaXml)).Attribute("FieldRef"), // FieldRef means this is a dependent lookup
                        Step = fld.GetFieldProvisioningStep(parser)
                    })
                    .Where(fldData => fldData.Step == step) // Only include fields related to the current step
                    .OrderBy(fldData => fldData.FieldRef) // Ensure fields having fieldRef are handled after. This ensure lookups are created before dependent lookups
                    .Select(fldData => fldData.Field)
                    .ToArray();

                foreach (var field in fieldsToProcess)
                {
                    var fieldElement = XElement.Parse(parser.ParseXmlString(field.SchemaXml, "~sitecollection", "~site"));
                    if (fieldElement.Attribute("ID") == null)
                    {
                        scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field_schema_has_no_ID_attribute___0_, field.SchemaXml);
                        throw new Exception(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field_schema_has_no_ID_attribute___0_, field.SchemaXml));
                    }
                    var id = fieldElement.Attribute("ID").Value;
                    var internalName = fieldElement.Attribute("InternalName")?.Value;

                    currentFieldIndex++;
                    WriteMessage($"List Columns for list {listInfo.TemplateList.Title}|{internalName ?? id}|{currentFieldIndex}|{fieldsToProcess.Length}", ProvisioningMessageType.Progress);
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
                                    parser.AddToken(new FieldTitleToken(web, updatedField.InternalName, updatedField.Title));
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

        private void ProcessFieldRefs(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope, Web rootWeb, ListInfo listInfo)
        {
            if (listInfo.TemplateList.FieldRefs.Any())
            {
                var fieldsRefsToProcess = listInfo.TemplateList.FieldRefs.Select(fr => new
                {
                    FieldRef = fr,
                    TemplateField = template.SiteFields.FirstOrDefault(tf => (Guid)XElement.Parse(parser.ParseString(tf.SchemaXml)).Attribute("ID") == fr.Id)
                }).Where(frData =>
                    frData.TemplateField == null // Process fields refs if the target is not defined in the current template
                    || frData.TemplateField.GetFieldProvisioningStep(parser) == step // or process field ref only if the current step is matching
                ).Select(fr => fr.FieldRef).ToArray();

                var total = fieldsRefsToProcess.Length;

                var currentListIndex = 0;
                foreach (var fieldRef in fieldsRefsToProcess)
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

        private static void CheckContentTypes(Web web, ProvisioningTemplate template, PnPMonitoredScope scope, ListInstance templateList)
        {
            // Check for the presence of the references content types and throw an exception if not present or in template
            if (templateList.ContentTypesEnabled)
            {
                var existingCts = web.Context.LoadQuery(web.AvailableContentTypes);
                web.Context.ExecuteQueryRetry();
                foreach (var ct in templateList.ContentTypeBindings)
                {
                    var found = template.ContentTypes.Any(t => string.Equals(t.Id, ct.ContentTypeId, StringComparison.InvariantCultureIgnoreCase));
                    if (!found)
                    {
                        found = existingCts.Any(t => string.Equals(t.StringId, ct.ContentTypeId, StringComparison.InvariantCultureIgnoreCase));
                    }
                    if (!found)
                    {
                        scope.LogError("Referenced content type {0} not available in site or in template", ct.ContentTypeId);
                        throw new Exception($"Referenced content type {ct.ContentTypeId} not available in site or in template");
                    }
                }
            }
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

                // Fix the calendar recurrence
                if (viewType == ViewType.Calendar)
                {
                    viewType = ViewType.Calendar | ViewType.Recurrence;
                }

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

                var reader = viewElement.CreateReader();
                reader.MoveToContent();
                var viewInnerXml = reader.ReadInnerXml();

                var createdView = createdList.Views.Add(viewCI);
                createdView.ListViewXml = viewInnerXml;
                createdView.Update();
                createdView.EnsureProperties(v => v.Scope, v => v.JSLink, v => v.Title, v => v.Aggregations, v => v.MobileView, v => v.MobileDefaultView, v => v.ViewData);
                web.Context.ExecuteQueryRetry();

                if (urlHasValue)
                {
                    //restore original title
                    createdView.Title = viewTitle;
                    createdView.Update();
                }

                // ContentTypeID
                var contentTypeID = (string)viewElement.Attribute("ContentTypeID");
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
                var defaultViewForContentType = (string)viewElement.Attribute("DefaultViewForContentType");
                if (!string.IsNullOrEmpty(defaultViewForContentType) && bool.TryParse(defaultViewForContentType, out parsedDefaultViewForContentType))
                {
                    createdView.DefaultViewForContentType = parsedDefaultViewForContentType;
                    createdView.Update();
                }

                // Scope
                var scope = (string)viewElement.Attribute("Scope");
                var parsedScope = ViewScope.DefaultValue;
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
                if (aggregationsElement != null && aggregationsElement.HasElements)
                {
                    var fieldRefString = "";
                    foreach (var fieldRef in aggregationsElement.Descendants("FieldRef"))
                    {
                        fieldRefString += fieldRef.ToString();
                    }
                    if (createdView.Aggregations != fieldRefString)
                    {
                        createdView.Aggregations = fieldRefString;
                        createdView.Update();
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

                // View Data
                var viewDataElement = viewElement.Descendants("ViewData").FirstOrDefault();
                if (viewDataElement != null && viewDataElement.HasElements)
                {
                    var fieldRefString = "";
                    foreach (var fieldRef in viewDataElement.Descendants("FieldRef"))
                    {
                        fieldRefString += fieldRef.ToString();
                    }
                    if (createdView.ViewData != fieldRefString)
                    {
                        createdView.ViewData = fieldRefString;
                        createdView.Update();
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
            if (!string.IsNullOrEmpty(fieldRef.DisplayName)
                && (fieldRef.DisplayName != listField.Title || fieldRef.DisplayName.ContainsResourceToken())
                )
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
            if (!(siteList.BaseTemplate == (int)ListTemplateType.Contacts
                && (fieldRef.Name.Equals("LastNamePhonetic", StringComparison.InvariantCultureIgnoreCase)
                || fieldRef.Name.Equals("FirstNamePhonetic", StringComparison.InvariantCultureIgnoreCase)
                || fieldRef.Name.Equals("CompanyPhonetic", StringComparison.InvariantCultureIgnoreCase))))
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

            var fieldXml = parser.ParseXmlString(fieldElement.ToString(), "~sitecollection", "~site");
            if (IsFieldXmlValid(parser.ParseXmlString(originalFieldXml), parser, context))
            {
                var addOptions = listInfo.TemplateList.ContentTypesEnabled
                    ? AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToNoContentType
                    : AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToDefaultContentType;
                field = listInfo.SiteList.Fields.AddFieldAsXml(fieldXml, false, addOptions);
                listInfo.SiteList.Context.Load(field);
                listInfo.SiteList.Context.ExecuteQueryRetry();

                bool isDirty = false;
#if !SP2013
                if (originalFieldXml.ContainsResourceToken())
                {
                    var originalFieldElement = XElement.Parse(originalFieldXml);
                    var nameAttributeValue = (string)originalFieldElement.Attribute("DisplayName");
                    if (nameAttributeValue.ContainsResourceToken()
                        && field.TitleResource.SetUserResourceValue(nameAttributeValue, parser))
                    {
                        isDirty = true;
                    }
                    var descriptionAttributeValue = (string)originalFieldElement.Attribute("Description");
                    if (descriptionAttributeValue.ContainsResourceToken()
                        && field.DescriptionResource.SetUserResourceValue(descriptionAttributeValue, parser))
                    {
                        isDirty = true;
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
                    if (IsFieldXmlValid(parser.ParseXmlString(templateFieldElement.ToString()), parser, web.Context))
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
                        existingField.SchemaXml = parser.ParseXmlString(existingFieldElement.ToString(), "~sitecollection", "~site");
                        existingField.UpdateAndPushChanges(true);
                        web.Context.ExecuteQueryRetry();
#if !SP2013
                        bool isDirty = false;
                        if (originalFieldXml.ContainsResourceToken())
                        {
                            var originalFieldElement = XElement.Parse(originalFieldXml);
                            var nameAttributeValue = (string)originalFieldElement.Attribute("DisplayName");
                            if (nameAttributeValue.ContainsResourceToken()
                                && existingField.TitleResource.SetUserResourceValue(nameAttributeValue, parser)
                                )
                            {
                                isDirty = true;
                            }
                            var descriptionAttributeValue = (string)originalFieldElement.Attribute("Description");
                            if (descriptionAttributeValue.ContainsResourceToken()
                                && existingField.DescriptionResource.SetUserResourceValue(descriptionAttributeValue, parser)
                                )
                            {
                                isDirty = true;
                            }
                        }
                        if (isDirty)
                        {
                            existingField.Update();
                            web.Context.ExecuteQueryRetry();
                            field = existingField;
                        }
#endif
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
                    var fieldName = (string)existingFieldElement.Attribute("Name") ?? (string)existingFieldElement.Attribute("StaticName");
                    var warning = string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_Field__0____1___exists_in_list__2____3___but_is_of_different_type__Skipping_field_, fieldName, fieldId, listInfo.TemplateList.Title, listInfo.SiteList.Id);
                    scope.LogWarning(warning);
                    WriteMessage(warning, ProvisioningMessageType.Warning);
                }
            }
            return field;
        }

        private static XElement PrepareField(XElement fieldElement)
        {
            var listIdentifier = (string)fieldElement.Attribute("List");

            if (listIdentifier != null)
            {
                if (fieldElement.Attribute("RelationshipDeleteBehavior") != null)
                {
                    if (fieldElement.Attribute("RelationshipDeleteBehavior").Value.Equals("Restrict")
                        || fieldElement.Attribute("RelationshipDeleteBehavior").Value.Equals("Cascade"))
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
                l => l.DefaultDisplayFormUrl,
                l => l.DefaultEditFormUrl,
                l => l.DefaultNewFormUrl,
                l => l.IsApplicationList,
                l => l.Direction,
                l => l.ImageUrl,
                l => l.IrmExpire,
                l => l.IrmReject,
                l => l.IrmEnabled,
                l => l.ValidationFormula,
                l => l.ValidationMessage,
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
#if !ONPREMISES
, l => l.ListExperienceOptions
, l => l.ReadSecurity
, l => l.WriteSecurity
#endif
);
            web.Context.ExecuteQueryRetry();

            if (existingList.BaseTemplate == templateList.TemplateType)
            {
                var isDirty = false;

#if !SP2013
                string newUrl = UrlUtility.Combine(web.ServerRelativeUrl, templateList.Url);
                string oldUrl = existingList.RootFolder.ServerRelativeUrl;
                if (!newUrl.Equals(oldUrl, StringComparison.OrdinalIgnoreCase))
                {
                    Microsoft.SharePoint.Client.Folder folder = web.GetFolderByServerRelativeUrl(oldUrl);
                    folder.MoveTo(newUrl);
                    folder.Update();
                }
#endif

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
                isDirty |= existingList.Set(x => x.DocumentTemplateUrl, parser.ParseString(templateList.DocumentTemplate).NullIfEmpty(), false, false);
                isDirty |= existingList.Set(x => x.Description, parser.ParseString(templateList.Description), false, false);
                isDirty |= existingList.Set(x => x.Hidden, templateList.Hidden);
                isDirty |= existingList.Set(x => x.OnQuickLaunch, templateList.OnQuickLaunch);
                isDirty |= existingList.Set(x => x.DefaultDisplayFormUrl, parser.ParseString(templateList.DefaultDisplayFormUrl).NullIfEmpty(), false);
                isDirty |= existingList.Set(x => x.DefaultEditFormUrl, parser.ParseString(templateList.DefaultEditFormUrl).NullIfEmpty(), false);
                isDirty |= existingList.Set(x => x.DefaultNewFormUrl, parser.ParseString(templateList.DefaultNewFormUrl).NullIfEmpty(), false);

                if (existingList.Direction == "none" && templateList.Direction != ListReadingDirection.None)
                {
                    existingList.Direction = templateList.Direction == ListReadingDirection.None ? "none" : templateList.Direction == ListReadingDirection.RTL ? "rtl" : "ltr";
                    isDirty = true;
                }
                else if (existingList.Direction == "rtl" && templateList.Direction != ListReadingDirection.RTL)
                {
                    existingList.Direction = templateList.Direction == ListReadingDirection.None ? "none" : templateList.Direction == ListReadingDirection.RTL ? "rtl" : "ltr";
                    isDirty = true;
                }
                else if (existingList.Direction == "ltr" && templateList.Direction != ListReadingDirection.LTR)
                {
                    existingList.Direction = templateList.Direction == ListReadingDirection.None ? "none" : templateList.Direction == ListReadingDirection.RTL ? "rtl" : "ltr";
                    isDirty = true;
                }

                isDirty |= existingList.Set(x => x.ImageUrl, parser.ParseString(templateList.ImageUrl), false);
                isDirty |= existingList.Set(x => x.IsApplicationList, templateList.IsApplicationList);

#if !ONPREMISES
                if (existingList.ReadSecurity != (templateList.ReadSecurity == 0 ? 1 : templateList.ReadSecurity))
                {
                    // 0 or 1 [Default] = Read all items
                    // 2 = Read items that where created by the user
                    existingList.ReadSecurity = (templateList.ReadSecurity == 0 ? 1 : templateList.ReadSecurity);
                    isDirty = true;
                }
                if (existingList.WriteSecurity != (templateList.WriteSecurity == 0 ? 1 : templateList.WriteSecurity))
                {
                    // 0 or 1 [Default] = Create and edit all items
                    // 2 = Create items and edit items that where created by the user
                    // 4 = None
                    existingList.WriteSecurity = (templateList.WriteSecurity == 0 ? 1 : templateList.WriteSecurity);
                    isDirty = true;
                }
#endif

                isDirty |= existingList.Set(x => x.ValidationFormula, parser.ParseString(templateList.ValidationFormula), false);
                isDirty |= existingList.Set(x => x.ValidationMessage, parser.ParseString(templateList.ValidationMessage), false);
                isDirty |= existingList.Set(x => x.IrmExpire, templateList.IrmExpire);
                isDirty |= existingList.Set(x => x.IrmReject, templateList.IrmReject);

                if (existingList.BaseTemplate != (int)ListTemplateType.PictureLibrary && templateList.IRMSettings != null)
                {
                    isDirty |= existingList.Set(x => x.IrmEnabled, templateList.IRMSettings.Enabled);

                    existingList.EnsureProperties(l => l.InformationRightsManagementSettings);

                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.AllowPrint, templateList.IRMSettings.AllowPrint);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.AllowScript, templateList.IRMSettings.AllowScript);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.AllowWriteCopy, templateList.IRMSettings.AllowWriteCopy);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.DisableDocumentBrowserView, templateList.IRMSettings.DisableDocumentBrowserView);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.DocumentAccessExpireDays, templateList.IRMSettings.DocumentAccessExpireDays);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.DocumentLibraryProtectionExpireDate, DateTime.Now.AddDays(templateList.IRMSettings.DocumentLibraryProtectionExpiresInDays));
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.EnableDocumentAccessExpire, templateList.IRMSettings.EnableDocumentAccessExpire);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.EnableDocumentBrowserPublishingView, templateList.IRMSettings.EnableDocumentBrowserPublishingView);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.EnableGroupProtection, templateList.IRMSettings.EnableGroupProtection);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.EnableLicenseCacheExpire, templateList.IRMSettings.EnableLicenseCacheExpire);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.GroupName, parser.ParseString(templateList.IRMSettings.GroupName));
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.LicenseCacheExpireDays, templateList.IRMSettings.LicenseCacheExpireDays);
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.PolicyDescription, parser.ParseString(templateList.IRMSettings.PolicyDescription));
                    isDirty |= existingList.Set(x => x.InformationRightsManagementSettings.PolicyTitle, parser.ParseString(templateList.IRMSettings.PolicyTitle));
                }
                if (existingList.BaseTemplate != (int)ListTemplateType.Survey
                    && templateList.ContentTypesEnabled != existingList.ContentTypesEnabled)
                {
                    existingList.ContentTypesEnabled = templateList.ContentTypesEnabled;
                    isDirty = true;
                }
#if !ONPREMISES
                isDirty |= existingList.Set(x => x.ListExperienceOptions, (Microsoft.SharePoint.Client.ListExperience)Enum.Parse(typeof(Microsoft.SharePoint.Client.ListExperience), templateList.ListExperience.ToString()));

#endif
                if (existingList.BaseTemplate != (int)ListTemplateType.Survey
                    && existingList.BaseTemplate != (int)ListTemplateType.DocumentLibrary
                    && existingList.BaseTemplate != (int)ListTemplateType.PictureLibrary
                    && existingList.BaseTemplate != 850) // 850 = Pages library on publishing site
                {
                    // https://msdn.microsoft.com/EN-US/library/microsoft.sharepoint.splist.enableattachments.aspx
                    // The EnableAttachments property does not apply to any list that has a base type of Survey, DocumentLibrary or PictureLibrary.
                    // If you set this property to true for either type of list, it throws an SPException.
                    isDirty |= existingList.Set(x => x.EnableAttachments, templateList.EnableAttachments);
                }
                if (existingList.BaseTemplate != (int)ListTemplateType.DiscussionBoard)
                {
                    isDirty |= existingList.Set(x => x.EnableFolderCreation, templateList.EnableFolderCreation);
                }
#if !SP2013
                if (templateList.Title.ContainsResourceToken())
                {
                    if (existingList.TitleResource.SetUserResourceValue(templateList.Title, parser))
                    {
                        isDirty = true;
                    }
                }

                if (templateList.Description.ContainsResourceToken())
                {
                    if (existingList.DescriptionResource.SetUserResourceValue(templateList.Description, parser))
                    {
                        isDirty = true;
                    }
                }
#endif
                isDirty |= existingList.Set(x => x.EnableModeration, templateList.EnableModeration);
                isDirty |= existingList.Set(x => x.ForceCheckout, templateList.ForceCheckout);

                if (templateList.EnableVersioning)
                {
                    isDirty |= existingList.Set(x => x.EnableVersioning, templateList.EnableVersioning);

#if !SP2013
                    isDirty |= existingList.Set(x => x.MajorVersionLimit, templateList.MaxVersionLimit);

#endif
                    if (existingList.BaseType == BaseType.DocumentLibrary)
                    {
                        // Only supported on Document Libraries
                        isDirty |= existingList.Set(x => x.EnableMinorVersions, templateList.EnableMinorVersions);
                        isDirty |= existingList.Set(x => x.DraftVersionVisibility, (DraftVisibilityType)templateList.DraftVersionVisibility);

                        if (templateList.EnableMinorVersions)
                        {
                            isDirty |= existingList.Set(x => x.MajorWithMinorVersionsLimit, templateList.MinorVersionLimit);

                            if (DraftVisibilityType.Approver == (DraftVisibilityType)templateList.DraftVersionVisibility)
                            {
                                if (templateList.EnableModeration)
                                {
                                    isDirty |= existingList.Set(x => x.DraftVersionVisibility, (DraftVisibilityType)templateList.DraftVersionVisibility);
                                }
                            }
                            else
                            {
                                isDirty |= existingList.Set(x => x.DraftVersionVisibility, (DraftVisibilityType)templateList.DraftVersionVisibility);
                            }
                        }
                    }
                }
                else
                {
                    isDirty |= existingList.Set(x => x.EnableVersioning, templateList.EnableVersioning);
                }

                isDirty |= existingList.Set(x => x.NoCrawl, templateList.NoCrawl);

                if (isDirty)
                {
                    existingList.Update();
                    web.Context.ExecuteQueryRetry();
                    isDirty = false;
                }

#if !ONPREMISES
                // Process list webhooks
                if (templateList.Webhooks.Any())
                {
                    foreach (var webhook in templateList.Webhooks)
                    {
                        AddOrUpdateListWebHook(existingList, webhook, scope, parser, true);
                    }
                }
#endif

                #region UserCustomActions

                isDirty |= UpdateCustomActions(web, existingList, templateList, parser, scope, isNoScriptSite);

                #endregion UserCustomActions

                if (isDirty)
                {
                    existingList.Update();
                    web.Context.ExecuteQueryRetry();
                    isDirty = false;
                }

                if (existingList.ContentTypesEnabled)
                {
                    ConfigureContentTypes(web, existingList, templateList, false, scope);
                }
                if (templateList.Security != null)
                {
                    existingList.SetSecurity(parser, templateList.Security);
                }
                return Tuple.Create(existingList, parser);
            }
            else
            {
                var warning = string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_List__0____1____2___exists_but_is_of_a_different_type__Skipping_list_, templateList.Title, templateList.Url, existingList.Id);
                scope.LogWarning(warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
                return null;
            }
        }

        private static bool UpdateCustomActions(Web web, List existingList, ListInstance templateList, TokenParser parser, PnPMonitoredScope scope, bool isNoScriptSite)
        {
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
                        return true;
                    }
                    else
                    {
                        var existingCustomAction = existingUserCustomActions.AsEnumerable().FirstOrDefault(uca => uca.Name == userCustomAction.Name);
                        if (existingCustomAction != null)
                        {
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
                            return true;
                        }
                    }
                }
            }
            else
            {
                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_SkipAddingOrUpdatingCustomActions);
            }

            return false;
        }

        private void ConfigureContentTypes(Web web, List list, ListInstance templateList, bool isNewList, PnPMonitoredScope scope)
        {
            var contentTypesToRemove = new List<ContentType>();

            if (isNewList)
            {
                if (templateList.RemoveExistingContentTypes && templateList.ContentTypeBindings.Count > 0)
                {
                    contentTypesToRemove.AddRange(list.ContentTypes);
                }
            }
            else
            {
                var existingContentTypes = list.ContentTypes;
                web.Context.Load(existingContentTypes, cts => cts.Include(ct => ct.StringId));
                web.Context.Load(existingContentTypes, cts => cts.Include(ct => ct.Id));
                web.Context.ExecuteQueryRetry();

                if (templateList.RemoveExistingContentTypes && existingContentTypes.Count > 0)
                {
                    var warning = $"You specified to remove existing content types for the list with url '{list.RootFolder.ServerRelativeUrl}'. We found a list with the same url in the site. In case of a list update we cannot remove existing content types as they can be in use by existing list items and/or documents.";
                    scope.LogWarning(warning);
                    WriteMessage(warning, ProvisioningMessageType.Warning);
                }
            }

            ContentType defaultContentType = null;
            IList<ContentType> contentTypesToShowInNewButton = new List<ContentType>();
            IList<ContentType> contentTypesToHideInNewButton = new List<ContentType>();

            foreach (var ctb in templateList.ContentTypeBindings)
            {
                var tempCT = web.GetContentTypeById(ctb.ContentTypeId, searchInSiteHierarchy: true);
                if (tempCT != null)
                {
                    ContentTypeId existingContentTypeId = list.ContentTypes.BestMatch(ctb.ContentTypeId);
                    bool contentTypeAlreadyExistsInList = existingContentTypeId != null && existingContentTypeId.GetParentIdValue().Equals(ctb.ContentTypeId, StringComparison.OrdinalIgnoreCase);
                    if (ctb.Remove)
                    {
                        if (contentTypeAlreadyExistsInList)
                        {
                            // Remove the content type
                            list.ContentTypes.GetById(existingContentTypeId.StringValue).DeleteObject();
                            list.Update();
                            web.Context.ExecuteQueryRetry();
                        }
                    }
                    else
                    {
                        ContentType listContentType;
                        if (contentTypeAlreadyExistsInList)
                        {
                            //Get the content type
                            listContentType = list.ContentTypes.GetById(existingContentTypeId.StringValue);
                        }
                        else
                        {
                            // Add the content type
                            listContentType = list.ContentTypes.AddExistingContentType(tempCT);
                        }
                        web.Context.ExecuteQueryRetry();

                        if (ctb.Default && defaultContentType == null)
                        {
                            defaultContentType = listContentType;
                        }

                        if (listContentType.GetIsAllowedInContentTypeOrder())
                        {
                            if (ctb.Hidden)
                            {
                                contentTypesToHideInNewButton.Add(listContentType);
                            }
                            else
                            {
                                contentTypesToShowInNewButton.Add(listContentType);
                            }
                        }
                    }
                }
            }

            // Effectively remove existing content types, if any
            foreach (var ct in contentTypesToRemove)
            {
                var shouldDelete = true;
                shouldDelete &= ((list.BaseTemplate != (int)ListTemplateType.DocumentLibrary
                    && list.BaseTemplate != 851)
                    || !ct.StringId.StartsWith(BuiltInContentTypeId.Folder + "00"));

                if (shouldDelete)
                {
                    ct.DeleteObject();
                    web.Context.ExecuteQueryRetry();
                }
            }

            //Content type order and visibility should be done after removing all pre-existing content types.
            //If the content type configuration matches the content type order on the list
            //a unique content type order is not required.
            list.ShowContentTypesInNewButton(contentTypesToShowInNewButton);
            list.HideContentTypesInNewButton(contentTypesToHideInNewButton);

            if (defaultContentType != null)
            {
                defaultContentType.EnsureProperty(ct => ct.Id);
                list.SetDefaultContentType(defaultContentType.Id);
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
            if (userCustomAction.ClientSideComponentId != null && userCustomAction.ClientSideComponentId != Guid.Empty)
            {
                newUserCustomAction.ClientSideComponentId = userCustomAction.ClientSideComponentId;
            }
            if (!string.IsNullOrEmpty(userCustomAction.ClientSideComponentProperties))
            {
                newUserCustomAction.ClientSideComponentProperties = parser.ParseString(userCustomAction.ClientSideComponentProperties);
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

        private Tuple<List, TokenParser> CreateList(Web web, ListInstance templateList, TokenParser parser, PnPMonitoredScope scope, bool isNoScriptSite = false)
        {
            List createdList;
            if (templateList.Url.Equals("SiteAssets") && templateList.TemplateType == (int)ListTemplateType.DocumentLibrary)
            {
                //Ensure that the Site Assets library is created using the out of the box creation mechanism
                //Site Assets that are created using the EnsureSiteAssetsLibrary method slightly differ from
                //default Document Libraries. See issue 512 (https://github.com/OfficeDev/PnP-Sites-Core/issues/512)
                //for details about the issue fixed by this approach.
                createdList = web.Lists.EnsureSiteAssetsLibrary();
                //Check that Title and Description have the correct values
                web.Context.Load(createdList, l => l.Title,
                                              l => l.Description,
                                              l => l.NoCrawl);
                web.Context.ExecuteQueryRetry();
                var isDirty = false;
                if (!string.Equals(createdList.Description, templateList.Description))
                {
                    createdList.Description = templateList.Description;
                    isDirty = true;
                }
                if (!string.Equals(createdList.Title, templateList.Title))
                {
                    createdList.Title = templateList.Title;
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
                ListCreationInformation listCreate =
                    new ListCreationInformation
                    {
                        Title = parser.ParseString(templateList.Title),
                        Description = parser.ParseString(templateList.Description),
                        Url = parser.ParseString(templateList.Url),
                        // the line of code below doesn't add the list to QuickLaunch
                        // the OnQuickLaunch property is re-set on the Created List object
                        QuickLaunchOption = templateList.OnQuickLaunch ? QuickLaunchOptions.On : QuickLaunchOptions.Off
                    };
#if !ONPREMISES
                if (templateList.TemplateFeatureID != Guid.Empty)
                {
                    Site site = ((ClientContext)web.Context).Site;
                    var listTemplates = site.GetCustomListTemplates(web);
                    web.Context.Load(listTemplates);
                    web.Context.ExecuteQueryRetry();
                    var template = listTemplates.SingleOrDefault(t => t.FeatureId == templateList.TemplateFeatureID);
                    if (template != null)
                    {
                        listCreate.ListTemplate = template;
                    }
                }
                if (listCreate.ListTemplate == null)
                {
                    listCreate.TemplateType = templateList.TemplateType;
                }
#else
                listCreate.TemplateType = templateList.TemplateType;
#endif
                createdList = web.Lists.Add(listCreate);
                createdList.Update();
            }
            web.Context.Load(createdList, l => l.BaseTemplate);
            web.Context.ExecuteQueryRetry();

#if !SP2013
            if (templateList.Title.ContainsResourceToken())
            {
                createdList.TitleResource.SetUserResourceValue(templateList.Title, parser);
            }
            if (templateList.Description.ContainsResourceToken())
            {
                createdList.DescriptionResource.SetUserResourceValue(templateList.Description, parser);
            }
#endif
            if (!String.IsNullOrEmpty(templateList.DocumentTemplate))
            {
                createdList.DocumentTemplateUrl = parser.ParseString(templateList.DocumentTemplate);
            }
            if (!string.IsNullOrEmpty(parser.ParseString(templateList.DefaultDisplayFormUrl)))
            {
                createdList.DefaultDisplayFormUrl = parser.ParseString(templateList.DefaultDisplayFormUrl);
            }
            if (!string.IsNullOrEmpty(parser.ParseString(templateList.DefaultEditFormUrl)))
            {
                createdList.DefaultEditFormUrl = parser.ParseString(templateList.DefaultEditFormUrl);
            }
            if (!string.IsNullOrEmpty(parser.ParseString(templateList.DefaultNewFormUrl)))
            {
                createdList.DefaultNewFormUrl = parser.ParseString(templateList.DefaultNewFormUrl);
            }
            createdList.Direction = templateList.Direction.ToString().ToLower();
            if (!string.IsNullOrEmpty(parser.ParseString(templateList.ImageUrl)))
            {
                createdList.ImageUrl = parser.ParseString(templateList.ImageUrl);
            }
            createdList.IrmExpire = templateList.IrmExpire;
            createdList.IrmReject = templateList.IrmReject;
            createdList.IsApplicationList = templateList.IsApplicationList;
#if !ONPREMISES
            if (templateList.ReadSecurity != default(int))
            {
                createdList.ReadSecurity = templateList.ReadSecurity;
            }
            if (templateList.WriteSecurity != default(int))
            {
                createdList.WriteSecurity = templateList.WriteSecurity;
            }
#endif
            if (!string.IsNullOrEmpty(parser.ParseString(templateList.ValidationFormula)))
            {
                createdList.ValidationFormula = parser.ParseString(templateList.ValidationFormula);
            }
            if (!string.IsNullOrEmpty(parser.ParseString(templateList.ValidationMessage)))
            {
                createdList.ValidationMessage = parser.ParseString(templateList.ValidationMessage);
            }
            if (createdList.BaseTemplate != (int)ListTemplateType.PictureLibrary && templateList.IRMSettings != null)
            {
                createdList.IrmEnabled = templateList.IRMSettings.Enabled;
                createdList.InformationRightsManagementSettings.AllowPrint = templateList.IRMSettings.AllowPrint;
                createdList.InformationRightsManagementSettings.AllowScript = templateList.IRMSettings.AllowScript;
                createdList.InformationRightsManagementSettings.AllowWriteCopy = templateList.IRMSettings.AllowWriteCopy;
                createdList.InformationRightsManagementSettings.DisableDocumentBrowserView = templateList.IRMSettings.DisableDocumentBrowserView;
                createdList.InformationRightsManagementSettings.DocumentAccessExpireDays = templateList.IRMSettings.DocumentAccessExpireDays;
                createdList.InformationRightsManagementSettings.DocumentLibraryProtectionExpireDate = DateTime.Now.AddDays(templateList.IRMSettings.DocumentLibraryProtectionExpiresInDays);
                createdList.InformationRightsManagementSettings.EnableDocumentAccessExpire = templateList.IRMSettings.EnableDocumentAccessExpire;
                createdList.InformationRightsManagementSettings.EnableDocumentBrowserPublishingView = templateList.IRMSettings.EnableDocumentBrowserPublishingView;
                createdList.InformationRightsManagementSettings.EnableGroupProtection = templateList.IRMSettings.EnableGroupProtection;
                createdList.InformationRightsManagementSettings.EnableLicenseCacheExpire = templateList.IRMSettings.EnableLicenseCacheExpire;
                if (!string.IsNullOrEmpty(parser.ParseString(templateList.IRMSettings.GroupName)))
                {
                    createdList.InformationRightsManagementSettings.GroupName = parser.ParseString(templateList.IRMSettings.GroupName);
                }
                if (!string.IsNullOrEmpty(parser.ParseString(templateList.IRMSettings.PolicyDescription)))
                {
                    createdList.InformationRightsManagementSettings.PolicyDescription = parser.ParseString(templateList.IRMSettings.PolicyDescription);
                }
                if (!string.IsNullOrEmpty(parser.ParseString(templateList.IRMSettings.PolicyTitle)))
                {
                    createdList.InformationRightsManagementSettings.PolicyTitle = parser.ParseString(templateList.IRMSettings.PolicyTitle);
                }
            }
#if !ONPREMISES
            createdList.ListExperienceOptions = (Microsoft.SharePoint.Client.ListExperience)Enum.Parse(typeof(Microsoft.SharePoint.Client.ListExperience), templateList.ListExperience.ToString());
#endif

            // EnableAttachments are not supported for DocumentLibraries, Survey and PictureLibraries
            // TODO: the user should be warned
            if (createdList.BaseTemplate != (int)ListTemplateType.DocumentLibrary
                && createdList.BaseTemplate != (int)ListTemplateType.Survey
                && createdList.BaseTemplate != (int)ListTemplateType.PictureLibrary)
            {
                createdList.EnableAttachments = templateList.EnableAttachments;
            }

            createdList.EnableModeration = templateList.EnableModeration;
            createdList.ForceCheckout = templateList.ForceCheckout;

            // Done for all other lists than for Survey - With Surveys versioning configuration will cause an exception
            if (createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                createdList.EnableVersioning = templateList.EnableVersioning;
                if (templateList.EnableVersioning)
                {
#if !SP2013
                    createdList.MajorVersionLimit = templateList.MaxVersionLimit;
#endif
                    // DraftVisibilityType.Approver is available only when the EnableModeration option of the list is true
                    if (DraftVisibilityType.Approver
                        == (DraftVisibilityType)templateList.DraftVersionVisibility)
                    {
                        if (templateList.EnableModeration)
                        {
                            createdList.DraftVersionVisibility =
                                (DraftVisibilityType)templateList.DraftVersionVisibility;
                        }
                        else
                        {
                            var warning = string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_DraftVersionVisibility_not_applied_to_list_0_because_EnableModeration_is_not_set_to_true, templateList.Url);
                            scope.LogWarning(warning);
                            WriteMessage(warning, ProvisioningMessageType.Warning);
                        }
                    }
                    else
                    {
                        createdList.DraftVersionVisibility = (DraftVisibilityType)templateList.DraftVersionVisibility;
                    }

                    if (createdList.BaseTemplate == (int)ListTemplateType.DocumentLibrary)
                    {
                        // Only supported on Document Libraries
                        createdList.EnableMinorVersions = templateList.EnableMinorVersions;
                        createdList.DraftVersionVisibility = (DraftVisibilityType)templateList.DraftVersionVisibility;

                        if (templateList.EnableMinorVersions)
                        {
                            createdList.MajorWithMinorVersionsLimit = templateList.MinorVersionLimit; // Set only if enabled, otherwise you'll get exception due setting value to zero.
                        }
                    }
                }
            }

            createdList.OnQuickLaunch = templateList.OnQuickLaunch;
            if (createdList.BaseTemplate != (int)ListTemplateType.DiscussionBoard
                && createdList.BaseTemplate != (int)ListTemplateType.Events)
            {
                createdList.EnableFolderCreation = templateList.EnableFolderCreation;
            }
            createdList.Hidden = templateList.Hidden;

            if (createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                createdList.ContentTypesEnabled = templateList.ContentTypesEnabled;
            }

            createdList.NoCrawl = templateList.NoCrawl;

            createdList.Update();

            web.Context.Load(createdList.Views);
            web.Context.Load(createdList, l => l.Id);
            web.Context.Load(createdList, l => l.RootFolder.ServerRelativeUrl);
            web.Context.Load(createdList.ContentTypes);
            web.Context.ExecuteQueryRetry();

            if (createdList.BaseTemplate != (int)ListTemplateType.Survey)
            {
                ConfigureContentTypes(web, createdList, templateList, true, scope);
            }

            // Add any custom action
            if (templateList.UserCustomActions.Any())
            {
                if (!isNoScriptSite)
                {
                    foreach (var userCustomAction in templateList.UserCustomActions)
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

#if !ONPREMISES
            // Process list webhooks
            if (templateList.Webhooks.Any())
            {
                foreach (var webhook in templateList.Webhooks)
                {
                    AddOrUpdateListWebHook(createdList, webhook, scope, parser);
                }
            }
#endif
            if (templateList.Security != null)
            {
                createdList.SetSecurity(parser, templateList.Security);
            }
            return Tuple.Create(createdList, parser);
        }

#if !ONPREMISES

        private void AddOrUpdateListWebHook(List list, Webhook webhook, PnPMonitoredScope scope, TokenParser parser, bool isListUpdate = false)
        {
            var webhookServerNotificationUrl = parser.ParseString(webhook.ServerNotificationUrl);
            if (webhook.ExpiresInDays > 0)
            {
                try
                {
                    // for a new list immediately add the webhook
                    if (!isListUpdate)
                    {
                        var webhookSubscription = list.AddWebhookSubscription(webhookServerNotificationUrl, DateTime.Now.AddDays(webhook.ExpiresInDays));
                    }
                    // for existing lists add a new webhook or update existing webhook
                    else
                    {
                        // get the webhooks defined on the list
                        var addedWebhooks = Task.Run(() => list.GetWebhookSubscriptionsAsync()).GetAwaiter().GetResult();

                        var existingWebhook = addedWebhooks.Where(p => p.NotificationUrl.Equals(webhookServerNotificationUrl, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                        if (existingWebhook != null)
                        {
                            // refresh the expiration date of the existing webhook
                            existingWebhook.ExpirationDateTime = DateTime.Now.AddDays(webhook.ExpiresInDays);
                            // update the existing webhook
                            list.UpdateWebhookSubscription(existingWebhook);
                        }
                        else
                        {
                            // add as new webhook
                            var webhookSubscription = list.AddWebhookSubscription(webhookServerNotificationUrl, DateTime.Now.AddDays(webhook.ExpiresInDays));
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Eat all webhook exceptions, we don't want to stop the provisioning flow is an exported file happended to have a reference to a stale webhook
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstances_Webhook_Error, ex.Message);
                }
            }
            else
            {
                list.EnsureProperty(l => l.Title);
                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstances_SkipExpiredWebHook, webhookServerNotificationUrl, list.Title);
            }
        }

#endif

        private void CreateFolderInList(Microsoft.SharePoint.Client.Folder parentFolder, Model.Folder folder, TokenParser parser, PnPMonitoredScope scope)
        {
            // Determine the folder name, parsing any token
            String targetFolderName = parser.ParseString(folder.Name);

            // Check if the folder already exists
            if (parentFolder.FolderExists(targetFolderName))
            {
                // Log a warning if the folder already exists
                var warningFolderAlreadyExists = String.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_FolderAlreadyExists, targetFolderName, parentFolder.ServerRelativeUrl);
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

                // Handle current folder property bags
                if (folder.PropertyBagEntries != null && folder.PropertyBagEntries.Count > 0)
                {
                    foreach (var p in folder.PropertyBagEntries)
                    {
                        currentFolder.Properties[p.Key] = parser.ParseString(p.Value);
                    }
                    currentFolder.Update();
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
                // Check if this is not a noscript site as we're not allowed to update some properties
                bool isNoScriptSite = web.IsNoScriptSite();

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
#if !SP2013
                        l => l.MajorVersionLimit,
                        l => l.MajorWithMinorVersionsLimit,
#endif
                        l => l.DraftVersionVisibility,
                        l => l.DefaultDisplayFormUrl,
                        l => l.DefaultEditFormUrl,
                        l => l.ImageUrl,
                        l => l.DefaultNewFormUrl,
                        l => l.Direction,
                        l => l.IrmExpire,
                        l => l.IrmReject,
                        l => l.IrmEnabled,
                        l => l.IsApplicationList,
                        l => l.ValidationFormula,
                        l => l.ValidationMessage,
                        l => l.DocumentTemplateUrl,
                        l => l.NoCrawl,
#if !ONPREMISES
                        l => l.ListExperienceOptions,
                        l => l.ReadSecurity,
                        l => l.WriteSecurity,
#endif
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
                    WriteMessage($"List|{siteList.Title}|{listCount}|{listsToProcess.Length}", ProvisioningMessageType.Progress);
                    ListInstance baseTemplateList = null;
                    if (creationInfo.BaseTemplate != null)
                    {
                        // Check if we need to skip this list...if so let's do it before we gather all the other information for this list...improves performance
                        var index = creationInfo.BaseTemplate.Lists.FindIndex(f => f.Url.Equals(siteList.RootFolder.ServerRelativeUrl.Substring(serverRelativeUrl.Length + 1))
                                                                                   && f.TemplateType.Equals(siteList.BaseTemplate));
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
                        DefaultDisplayFormUrl = Tokenize(siteList.DefaultDisplayFormUrl, web.Url),
                        DefaultEditFormUrl = Tokenize(siteList.DefaultEditFormUrl, web.Url),
                        DefaultNewFormUrl = Tokenize(siteList.DefaultNewFormUrl, web.Url),
                        Direction = string.Equals(siteList.Direction, "none", StringComparison.OrdinalIgnoreCase) ? ListReadingDirection.None : string.Equals(siteList.Direction, "rtl", StringComparison.OrdinalIgnoreCase) ? ListReadingDirection.RTL : ListReadingDirection.LTR,
                        ImageUrl = Tokenize(siteList.ImageUrl, web.Url),
                        IrmExpire = siteList.IrmExpire,
                        IrmReject = siteList.IrmReject,
                        IsApplicationList = siteList.IsApplicationList,
                        ValidationFormula = siteList.ValidationFormula,
                        ValidationMessage = siteList.ValidationMessage,
                        EnableModeration = siteList.EnableModeration,
                        NoCrawl = siteList.NoCrawl,
#if !ONPREMISES
                        ListExperience = (Model.ListExperience)Enum.Parse(typeof(Model.ListExperience), siteList.ListExperienceOptions.ToString()),
                        ReadSecurity = siteList.ReadSecurity,
                        WriteSecurity = siteList.WriteSecurity,
#endif
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

                    if (siteList.BaseTemplate != (int)ListTemplateType.PictureLibrary)
                    {
                        siteList.EnsureProperties(l => l.InformationRightsManagementSettings);
                    }

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

                    list = ExtractViews(web, siteList, list);

                    list = ExtractFields(web, siteList, contentTypeFields, list, allLists, creationInfo, template);

                    list = ExtractUserCustomActions(web, siteList, list, creationInfo, template);

#if !ONPREMISES
                    list = ExtractWebhooks(siteList, list);
#endif

                    list.Security = siteList.GetSecurity();

                    list = ExtractInformationRightsManagement(web, siteList, list, creationInfo, template);

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

#if !ONPREMISES

        private static ListInstance ExtractWebhooks(List siteList, ListInstance list)
        {
            var addedWebhooks = Task.Run(() => siteList.GetWebhookSubscriptionsAsync()).GetAwaiter().GetResult();

            if (addedWebhooks.Any())
            {
                foreach (var webhook in addedWebhooks)
                {
                    list.Webhooks.Add(new Webhook()
                    {
                        ExpiresInDays = webhook.ExpirationDateTime.Subtract(DateTime.Now).Days + 1,
                        ServerNotificationUrl = webhook.NotificationUrl,
                    });
                }
            }

            return list;
        }

#endif

        private static ListInstance ExtractViews(Web web, List siteList, ListInstance list)
        {
            foreach (var view in siteList.Views.AsEnumerable().Where(view => !view.Hidden && view.ListViewXml != null))
            {
                var schemaElement = XElement.Parse(view.ListViewXml);

                // exclude survey and events list as they dont support jsLink customizations
                if (siteList.BaseTemplate != (int)ListTemplateType.Survey && siteList.BaseTemplate != (int)ListTemplateType.Events)
                {
                    var currentView = siteList.GetViewById(view.Id);

                    Microsoft.SharePoint.Client.File viewPage = web.GetFileByServerRelativeUrl(currentView.ServerRelativeUrl);
                    Microsoft.SharePoint.Client.WebParts.LimitedWebPartManager limitedWebPartManager = viewPage.GetLimitedWebPartManager(Microsoft.SharePoint.Client.WebParts.PersonalizationScope.Shared);
                    web.Context.Load(limitedWebPartManager.WebParts);
                    web.Context.ExecuteQueryRetry();

                    if (limitedWebPartManager.WebParts.Count > 0)
                    {
                        var webPart = limitedWebPartManager.WebParts.FirstOrDefault();
                        web.Context.Load(webPart.WebPart.Properties);
                        web.Context.ExecuteQueryRetry();

                        if (webPart.WebPart.Properties.FieldValues.ContainsKey("JSLink"))
                        {
                            var jsLinkValue = webPart.WebPart.Properties["JSLink"];

                            var jsLinkElement = schemaElement.Descendants("JSLink").FirstOrDefault();
                            if (jsLinkElement != null && jsLinkValue != null)
                            {
                                jsLinkElement.Value = Convert.ToString(jsLinkValue);
                            }
                        }
                    }
                }

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
                        if (field.InternalName == "Editor"
                            || field.InternalName == "Author"
                            || field.InternalName == "Title"
                            || field.InternalName == "ID"
                            || field.InternalName == "Created"
                            || field.InternalName == "Modified"
                            || field.InternalName == "Attachments"
                            || field.InternalName == "_UIVersionString"
                            || field.InternalName == "DocIcon"
                            || field.InternalName == "LinkTitleNoMenu"
                            || field.InternalName == "LinkTitle"
                            || field.InternalName == "Edit"
                            || field.InternalName == "AppAuthor"
                            || field.InternalName == "AppEditor"
                            || field.InternalName == "ContentType"
                            || field.InternalName == "ItemChildCount"
                            || field.InternalName == "FolderChildCount"
                            || field.InternalName == "LinkFilenameNoMenu"
                            || field.InternalName == "LinkFilename"
                            || field.InternalName == "_CopySource"
                            || field.InternalName == "ParentVersionString"
                            || field.InternalName == "ParentLeafName"
                            || field.InternalName == "_CheckinComment"
                            || field.InternalName == "FileLeafRef"
                            || field.InternalName == "FileSizeDisplay"
                            || field.InternalName == "Preview"
                            || field.InternalName == "ThumbnailOnForm"
                            || field.InternalName == "CheckoutUser"
                            || field.InternalName == "Modified_x0020_By"
                            || field.InternalName == "Created_x0020_By"
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

        private static ListInstance ExtractInformationRightsManagement(Web web, List siteList, ListInstance list, ProvisioningTemplateCreationInformation creationInfo, ProvisioningTemplate template)
        {
            if (siteList.BaseTemplate != (int)ListTemplateType.PictureLibrary && siteList.IrmEnabled)
            {
                list.IRMSettings = new IRMSettings();
                list.IRMSettings.Enabled = siteList.IrmEnabled;
                list.IrmExpire = siteList.IrmExpire;
                list.IrmReject = siteList.IrmReject;

                list.IRMSettings.AllowPrint = siteList.InformationRightsManagementSettings.AllowPrint;
                list.IRMSettings.AllowScript = siteList.InformationRightsManagementSettings.AllowScript;
                list.IRMSettings.AllowWriteCopy = siteList.InformationRightsManagementSettings.AllowWriteCopy;
                list.IRMSettings.DisableDocumentBrowserView = siteList.InformationRightsManagementSettings.DisableDocumentBrowserView;
                list.IRMSettings.DocumentAccessExpireDays = siteList.InformationRightsManagementSettings.DocumentAccessExpireDays;
                list.IRMSettings.DocumentLibraryProtectionExpiresInDays = (Int32)siteList.InformationRightsManagementSettings.DocumentLibraryProtectionExpireDate.Subtract(DateTime.Now).TotalDays;
                list.IRMSettings.EnableDocumentAccessExpire = siteList.InformationRightsManagementSettings.EnableDocumentAccessExpire;
                list.IRMSettings.EnableDocumentBrowserPublishingView = siteList.InformationRightsManagementSettings.EnableDocumentBrowserPublishingView;
                list.IRMSettings.EnableGroupProtection = siteList.InformationRightsManagementSettings.EnableGroupProtection;
                list.IRMSettings.EnableLicenseCacheExpire = siteList.InformationRightsManagementSettings.EnableLicenseCacheExpire;
                list.IRMSettings.GroupName = siteList.InformationRightsManagementSettings.GroupName;
                list.IRMSettings.LicenseCacheExpireDays = siteList.InformationRightsManagementSettings.LicenseCacheExpireDays;
                list.IRMSettings.PolicyDescription = siteList.InformationRightsManagementSettings.PolicyDescription;
                list.IRMSettings.PolicyTitle = siteList.InformationRightsManagementSettings.PolicyTitle;
            }

            return (list);
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
                customAction.ClientSideComponentId = userCustomAction.ClientSideComponentId;
                customAction.ClientSideComponentProperties = userCustomAction.ClientSideComponentProperties;
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

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
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