using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201508;
using ContentType = OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    internal class XMLPnPSchemaV201508Formatter :
        IXMLSchemaFormatter, ITemplateFormatter
    {
        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        string IXMLSchemaFormatter.NamespaceUri
        {
#pragma warning disable 618
            get { return (XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_08); }
#pragma warning restore 618
        }

        string IXMLSchemaFormatter.NamespacePrefix
        {
            get { return (XMLConstants.PROVISIONING_SCHEMA_PREFIX); }
        }

        public bool IsValid(Stream template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            // Load the template into an XDocument
            XDocument xml = XDocument.Load(template);

            // Load the XSD embedded resource
            Stream stream = typeof(XMLPnPSchemaV201508Formatter)
                .Assembly
                .GetManifestResourceStream("OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2015-08.xsd");

            // Prepare the XML Schema Set
            XmlSchemaSet schemas = new XmlSchemaSet();
#pragma warning disable 618
            schemas.Add(XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_08,
#pragma warning restore 618
                new XmlTextReader(stream));

            Boolean result = true;
            xml.Validate(schemas, (o, e) =>
            {
                Diagnostics.Log.Error(e.Exception, "SchemaFormatter", "Template is not valid: {0}", e.Message);
                result = false;
            });

            return (result);
        }

        Stream ITemplateFormatter.ToFormattedTemplate(Model.ProvisioningTemplate template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            V201508.ProvisioningTemplate result = new V201508.ProvisioningTemplate();

            V201508.Provisioning wrappedResult = new V201508.Provisioning();
            wrappedResult.Preferences = new V201508.Preferences
            {
                Generator = this.GetType().Assembly.FullName
            };
            wrappedResult.Templates = new V201508.Templates[] {
                new V201508.Templates
                {
                    ID = $"CONTAINER-{template.Id}",
                    ProvisioningTemplate = new V201508.ProvisioningTemplate[]
                    {
                        result
                    }
                }
            };

            #region Basic Properties

            // Translate basic properties
            result.ID = template.Id;
            result.Version = (Decimal)template.Version;
            result.VersionSpecified = true;
            result.SitePolicy = template.SitePolicy;
            result.ImagePreviewUrl = template.ImagePreviewUrl;
            result.DisplayName = template.DisplayName;
            result.Description = template.Description;

            if (template.Properties != null && template.Properties.Any())
            {
                result.Properties =
                    (from p in template.Properties
                     select new V201508.StringDictionaryItem
                     {
                         Key = p.Key,
                         Value = p.Value,
                     }).ToArray();
            }
            else
            {
                result.Properties = null;
            }

            #endregion

            #region Property Bag

            // Translate PropertyBagEntries, if any
            if (template.PropertyBagEntries != null && template.PropertyBagEntries.Count > 0)
            {
                result.PropertyBagEntries =
                    (from bag in template.PropertyBagEntries
                     select new V201508.PropertyBagEntry()
                     {
                         Key = bag.Key,
                         Value = bag.Value,
                         Indexed = bag.Indexed,
                         Overwrite = bag.Overwrite,
                         OverwriteSpecified = true,
                     }).ToArray();
            }
            else
            {
                result.PropertyBagEntries = null;
            }

            #endregion

            #region Regional Settings

            if (template.RegionalSettings != null)
            {
                result.RegionalSettings = new V201508.RegionalSettings()
                {
                    AdjustHijriDays = template.RegionalSettings.AdjustHijriDays,
                    AdjustHijriDaysSpecified = true,
                    AlternateCalendarType = template.RegionalSettings.AlternateCalendarType.FromTemplateToSchemaCalendarTypeV201508(),
                    AlternateCalendarTypeSpecified = true,
                    CalendarType = template.RegionalSettings.CalendarType.FromTemplateToSchemaCalendarTypeV201508(),
                    CalendarTypeSpecified = true,
                    Collation = template.RegionalSettings.Collation,
                    CollationSpecified = true,
                    FirstDayOfWeek = (V201508.DayOfWeek)Enum.Parse(typeof(V201508.DayOfWeek), template.RegionalSettings.FirstDayOfWeek.ToString()),
                    FirstDayOfWeekSpecified = true,
                    FirstWeekOfYear = template.RegionalSettings.FirstWeekOfYear,
                    FirstWeekOfYearSpecified = true,
                    LocaleId = template.RegionalSettings.LocaleId,
                    LocaleIdSpecified = true,
                    ShowWeeks = template.RegionalSettings.ShowWeeks,
                    ShowWeeksSpecified = true,
                    Time24 = template.RegionalSettings.Time24,
                    Time24Specified = true,
                    TimeZone = template.RegionalSettings.TimeZone.ToString(),
                    WorkDayEndHour = template.RegionalSettings.WorkDayEndHour.FromTemplateToSchemaWorkHourV201508(),
                    WorkDayEndHourSpecified = true,
                    WorkDays = template.RegionalSettings.WorkDays,
                    WorkDaysSpecified = true,
                    WorkDayStartHour = template.RegionalSettings.WorkDayStartHour.FromTemplateToSchemaWorkHourV201508(),
                    WorkDayStartHourSpecified = true,
                };
            }
            else
            {
                result.RegionalSettings = null;
            }

            #endregion

            #region Supported UI Languages

            if (template.SupportedUILanguages != null && template.SupportedUILanguages.Count > 0)
            {
                result.SupportedUILanguages =
                    (from l in template.SupportedUILanguages
                     select new V201508.SupportedUILanguagesSupportedUILanguage
                     {
                         LCID = l.LCID,
                     }).ToArray();
            }
            else
            {
                result.SupportedUILanguages = null;
            }

            #endregion

            #region Audit Settings

            if (template.AuditSettings != null)
            {
                result.AuditSettings = new V201508.AuditSettings
                {
                    AuditLogTrimmingRetention = template.AuditSettings.AuditLogTrimmingRetention,
                    AuditLogTrimmingRetentionSpecified = true,
                    TrimAuditLog = template.AuditSettings.TrimAuditLog,
                    TrimAuditLogSpecified = true,
                    Audit = template.AuditSettings.AuditFlags.FromTemplateToSchemaAuditsV201508(),
                };
            }
            else
            {
                result.AuditSettings = null;
            }

            #endregion

            #region Security

            // Translate Security configuration, if any
            if (template.Security != null)
            {
                result.Security = new V201508.Security();

                if (template.Security.AdditionalAdministrators != null && template.Security.AdditionalAdministrators.Count > 0)
                {
                    result.Security.AdditionalAdministrators =
                        (from user in template.Security.AdditionalAdministrators
                         select new V201508.User
                         {
                             Name = user.Name,
                         }).ToArray();
                }
                else
                {
                    result.Security.AdditionalAdministrators = null;
                }

                if (template.Security.AdditionalOwners != null && template.Security.AdditionalOwners.Count > 0)
                {
                    result.Security.AdditionalOwners =
                        (from user in template.Security.AdditionalOwners
                         select new V201508.User
                         {
                             Name = user.Name,
                         }).ToArray();
                }
                else
                {
                    result.Security.AdditionalOwners = null;
                }

                if (template.Security.AdditionalMembers != null && template.Security.AdditionalMembers.Count > 0)
                {
                    result.Security.AdditionalMembers =
                        (from user in template.Security.AdditionalMembers
                         select new V201508.User
                         {
                             Name = user.Name,
                         }).ToArray();
                }
                else
                {
                    result.Security.AdditionalMembers = null;
                }

                if (template.Security.AdditionalVisitors != null && template.Security.AdditionalVisitors.Count > 0)
                {
                    result.Security.AdditionalVisitors =
                        (from user in template.Security.AdditionalVisitors
                         select new V201508.User
                         {
                             Name = user.Name,
                         }).ToArray();
                }
                else
                {
                    result.Security.AdditionalVisitors = null;
                }

                if (template.Security.SiteGroups != null && template.Security.SiteGroups.Count > 0)
                {
                    result.Security.SiteGroups =
                        (from g in template.Security.SiteGroups
                         select new V201508.SiteGroup
                         {
                             AllowMembersEditMembership = g.AllowMembersEditMembership,
                             AllowMembersEditMembershipSpecified = true,
                             AllowRequestToJoinLeave = g.AllowRequestToJoinLeave,
                             AllowRequestToJoinLeaveSpecified = true,
                             AutoAcceptRequestToJoinLeave = g.AutoAcceptRequestToJoinLeave,
                             AutoAcceptRequestToJoinLeaveSpecified = true,
                             Description = g.Description,
                             Members = g.Members.Any() ? (from m in g.Members
                                                          select new V201508.User
                                                          {
                                                              Name = m.Name,
                                                          }).ToArray() : null,
                             OnlyAllowMembersViewMembership = g.OnlyAllowMembersViewMembership,
                             OnlyAllowMembersViewMembershipSpecified = true,
                             Owner = g.Owner,
                             RequestToJoinLeaveEmailSetting = g.RequestToJoinLeaveEmailSetting,
                             Title = g.Title,
                         }).ToArray();
                }
                else
                {
                    result.Security.SiteGroups = null;
                }

                result.Security.Permissions = new SecurityPermissions();
                if (template.Security.SiteSecurityPermissions != null)
                {
                    if (template.Security.SiteSecurityPermissions.RoleAssignments != null && template.Security.SiteSecurityPermissions.RoleAssignments.Count > 0)
                    {
                        result.Security.Permissions.RoleAssignments =
                            (from ra in template.Security.SiteSecurityPermissions.RoleAssignments
                             select new V201508.RoleAssignment
                             {
                                 Principal = ra.Principal,
                                 RoleDefinition = ra.RoleDefinition,
                             }).ToArray();
                    }
                    else
                    {
                        result.Security.Permissions.RoleAssignments = null;
                    }
                    if (template.Security.SiteSecurityPermissions.RoleDefinitions != null && template.Security.SiteSecurityPermissions.RoleDefinitions.Count > 0)
                    {
                        result.Security.Permissions.RoleDefinitions =
                            (from rd in template.Security.SiteSecurityPermissions.RoleDefinitions
                             select new V201508.RoleDefinition
                             {
                                 Description = rd.Description,
                                 Name = rd.Name,
                                 Permissions =
                                    (from p in rd.Permissions
                                     select (RoleDefinitionPermission)Enum.Parse(typeof(RoleDefinitionPermission), p.ToString())).ToArray(),
                             }).ToArray();
                    }
                    else
                    {
                        result.Security.Permissions.RoleDefinitions = null;
                    }
                }
            }

            #endregion

            #region Site Columns

            // Translate Site Columns (Fields), if any
            if (template.SiteFields != null && template.SiteFields.Count > 0)
            {
                result.SiteFields = new V201508.ProvisioningTemplateSiteFields
                {
                    Any =
                        (from field in template.SiteFields
                         select field.SchemaXml.ToXmlElement()).ToArray(),
                };
            }
            else
            {
                result.SiteFields = null;
            }

            #endregion

            #region Content Types

            // Translate ContentTypes, if any
            if (template.ContentTypes != null && template.ContentTypes.Count > 0)
            {
                result.ContentTypes =
                    (from ct in template.ContentTypes
                     select new V201508.ContentType
                     {
                         ID = ct.Id,
                         Description = ct.Description,
                         Group = ct.Group,
                         Name = ct.Name,
                         Sealed = ct.Sealed,
                         Hidden =  ct.Hidden,
                         ReadOnly = ct.ReadOnly,
                         FieldRefs = ct.FieldRefs.Count > 0 ?
                         (from fieldRef in ct.FieldRefs
                          select new V201508.ContentTypeFieldRef
                          {
                              Name = fieldRef.Name,
                              ID = fieldRef.Id.ToString(),
                              Hidden = fieldRef.Hidden,
                              Required = fieldRef.Required
                          }).ToArray() : null,
                         DocumentSetTemplate = ct.DocumentSetTemplate != null ?
                             new V201508.DocumentSetTemplate
                             {
                                 AllowedContentTypes = ct.DocumentSetTemplate.AllowedContentTypes.Count > 0 ?
                                     (from act in ct.DocumentSetTemplate.AllowedContentTypes
                                      select new DocumentSetTemplateAllowedContentType
                                      {
                                          ContentTypeID = act
                                      }).ToArray() : null,
                                 DefaultDocuments = ct.DocumentSetTemplate.DefaultDocuments.Count > 0 ?
                                     (from dd in ct.DocumentSetTemplate.DefaultDocuments
                                      select new DocumentSetTemplateDefaultDocument
                                      {
                                          ContentTypeID = dd.ContentTypeId,
                                          FileSourcePath = dd.FileSourcePath,
                                          Name = dd.Name,
                                      }).ToArray() : null,
                                 SharedFields = ct.DocumentSetTemplate.SharedFields.Count > 0 ?
                                     (from sf in ct.DocumentSetTemplate.SharedFields
                                      select new DocumentSetFieldRef
                                      {
                                          ID = sf.ToString(),
                                      }).ToArray() : null,
                                 WelcomePage = ct.DocumentSetTemplate.WelcomePage,
                                 WelcomePageFields = ct.DocumentSetTemplate.WelcomePageFields.Count > 0 ?
                                     (from wpf in ct.DocumentSetTemplate.WelcomePageFields
                                      select new DocumentSetFieldRef
                                      {
                                          ID = wpf.ToString(),
                                      }).ToArray() : null,
                             } : null,
                         DisplayFormUrl = ct.DisplayFormUrl,
                         EditFormUrl = ct.EditFormUrl,
                         NewFormUrl = ct.NewFormUrl,
                     }).ToArray();
            }
            else
            {
                result.ContentTypes = null;
            }

            #endregion

            #region List Instances

            // Translate Lists Instances, if any
            if (template.Lists != null && template.Lists.Count > 0)
            {
                result.Lists =
                    (from list in template.Lists
                     select new V201508.ListInstance
                     {
                         ContentTypesEnabled = list.ContentTypesEnabled,
                         Description = list.Description,
                         DocumentTemplate = list.DocumentTemplate,
                         EnableVersioning = list.EnableVersioning,
                         EnableMinorVersions = list.EnableMinorVersions,
                         EnableModeration = list.EnableModeration,
                         DraftVersionVisibility = list.DraftVersionVisibility,
                         DraftVersionVisibilitySpecified = true,
                         Hidden = list.Hidden,
                         MinorVersionLimit = list.MinorVersionLimit,
                         MinorVersionLimitSpecified = true,
                         MaxVersionLimit = list.MaxVersionLimit,
                         MaxVersionLimitSpecified = true,
                         OnQuickLaunch = list.OnQuickLaunch,
                         EnableAttachments = list.EnableAttachments,
                         EnableFolderCreation = list.EnableFolderCreation,
                         RemoveExistingContentTypes = list.RemoveExistingContentTypes,
                         TemplateFeatureID = list.TemplateFeatureID != Guid.Empty ? list.TemplateFeatureID.ToString() : null,
                         TemplateType = list.TemplateType,
                         Title = list.Title,
                         Url = list.Url,
                         ContentTypeBindings = list.ContentTypeBindings.Count > 0 ?
                            (from contentTypeBinding in list.ContentTypeBindings
                             select new V201508.ContentTypeBinding
                             {
                                 ContentTypeID = contentTypeBinding.ContentTypeId,
                                 Default = contentTypeBinding.Default,
                             }).ToArray() : null,
                         Views = list.Views.Count > 0 ?
                         new V201508.ListInstanceViews
                         {
                             Any =
                                (from view in list.Views
                                 select view.SchemaXml.ToXmlElement()).ToArray(),
                             RemoveExistingViews = list.RemoveExistingViews,
                         } : null,
                         Fields = list.Fields.Count > 0 ?
                         new V201508.ListInstanceFields
                         {
                             Any =
                             (from field in list.Fields
                              select field.SchemaXml.ToXmlElement()).ToArray(),
                         } : null,
                         FieldDefaults = list.FieldDefaults.Count > 0 ?
                            (from value in list.FieldDefaults
                             select new FieldDefault { FieldName = value.Key, Value = value.Value }).ToArray() : null,
                         FieldRefs = list.FieldRefs.Count > 0 ?
                         (from fieldRef in list.FieldRefs
                          select new V201508.ListInstanceFieldRef
                          {
                              Name = fieldRef.Name,
                              DisplayName = fieldRef.DisplayName,
                              Hidden = fieldRef.Hidden,
                              Required = fieldRef.Required,
                              ID = fieldRef.Id.ToString(),
                          }).ToArray() : null,
                         DataRows = list.DataRows.Count > 0 ?
                            (from dr in list.DataRows
                             select new ListInstanceDataRow
                             {
                                 DataValue = dr.Values.Count > 0 ?
                                    (from value in dr.Values
                                     select new DataValue { FieldName = value.Key, Value = value.Value }).ToArray() : null,
                                 Security = dr.Security.FromTemplateToSchemaObjectSecurityV201508()
                             }).ToArray() : null,
                         Security = list.Security.FromTemplateToSchemaObjectSecurityV201508(),
                     }).ToArray();
            }
            else
            {
                result.Lists = null;
            }

            #endregion

            #region Features

            // Translate Features, if any
            if (template.Features != null)
            {
                result.Features = new V201508.Features();

                // TODO: This nullability check could be useless, because
                // the SiteFeatures property is initialized in the Features
                // constructor
                if (template.Features.SiteFeatures != null && template.Features.SiteFeatures.Count > 0)
                {
                    result.Features.SiteFeatures =
                        (from feature in template.Features.SiteFeatures
                         select new V201508.Feature
                         {
                             ID = feature.Id.ToString(),
                             Deactivate = feature.Deactivate,
                         }).ToArray();
                }
                else
                {
                    result.Features.SiteFeatures = null;
                }

                // TODO: This nullability check could be useless, because
                // the WebFeatures property is initialized in the Features
                // constructor
                if (template.Features.WebFeatures != null && template.Features.WebFeatures.Count > 0)
                {
                    result.Features.WebFeatures =
                        (from feature in template.Features.WebFeatures
                         select new V201508.Feature
                         {
                             ID = feature.Id.ToString(),
                             Deactivate = feature.Deactivate,
                         }).ToArray();
                }
                else
                {
                    result.Features.WebFeatures = null;
                }
            }

            #endregion

            #region Custom Actions

            // Translate CustomActions, if any
            if (template.CustomActions != null && (template.CustomActions.SiteCustomActions.Any() || template.CustomActions.WebCustomActions.Any()))
            {
                result.CustomActions = new V201508.CustomActions();

                if (template.CustomActions.SiteCustomActions != null && template.CustomActions.SiteCustomActions.Count > 0)
                {
                    result.CustomActions.SiteCustomActions =
                        (from customAction in template.CustomActions.SiteCustomActions
                         select new V201508.CustomAction
                         {
                             CommandUIExtension = new CustomActionCommandUIExtension
                             {
                                 Any = customAction.CommandUIExtension != null ?
                                    (from x in customAction.CommandUIExtension.Elements() select x.ToXmlElement()).ToArray() : null,
                             },
                             Description = customAction.Description,
                             Enabled = customAction.Enabled,
                             Group = customAction.Group,
                             ImageUrl = customAction.ImageUrl,
                             Location = customAction.Location,
                             Name = customAction.Name,
#pragma warning disable 618
                             Rights = customAction.RightsValue,
#pragma warning restore 618
                             RightsSpecified = true,
                             ScriptBlock = customAction.ScriptBlock,
                             ScriptSrc = customAction.ScriptSrc,
                             Sequence = customAction.Sequence,
                             SequenceSpecified = true,
                             Title = customAction.Title,
                             Url = customAction.Url,
                         }).ToArray();
                }
                else
                {
                    result.CustomActions.SiteCustomActions = null;
                }

                if (template.CustomActions.WebCustomActions != null && template.CustomActions.WebCustomActions.Count > 0)
                {
                    result.CustomActions.WebCustomActions =
                        (from customAction in template.CustomActions.WebCustomActions
                         select new V201508.CustomAction
                         {
                             CommandUIExtension = new CustomActionCommandUIExtension
                             {
                                 Any = customAction.CommandUIExtension != null ?
                                    (from x in customAction.CommandUIExtension.Elements() select x.ToXmlElement()).ToArray() : null,
                             },
                             Description = customAction.Description,
                             Enabled = customAction.Enabled,
                             Group = customAction.Group,
                             ImageUrl = customAction.ImageUrl,
                             Location = customAction.Location,
                             Name = customAction.Name,
#pragma warning disable 618
                             Rights = customAction.RightsValue,
#pragma warning restore 618
                             RightsSpecified = true,
                             ScriptBlock = customAction.ScriptBlock,
                             ScriptSrc = customAction.ScriptSrc,
                             Sequence = customAction.Sequence,
                             SequenceSpecified = true,
                             Title = customAction.Title,
                             Url = customAction.Url,
                         }).ToArray();
                }
                else
                {
                    result.CustomActions.WebCustomActions = null;
                }
            }

            #endregion

            #region Files

            // Translate Files, if any
            if (template.Files != null && template.Files.Count > 0)
            {
                result.Files =
                    (from file in template.Files
                     select new V201508.File
                     {
                         Overwrite = file.Overwrite,
                         Src = file.Src,
                         Folder = file.Folder,
                         WebParts = file.WebParts.Count > 0 ?
                            (from wp in file.WebParts
                             select new V201508.WebPartPageWebPart
                             {
                                 Zone = wp.Zone,
                                 Order = (int)wp.Order,
                                 Contents = wp.Contents,
                                 Title = wp.Title,
                             }).ToArray() : null,
                         Properties = file.Properties != null && file.Properties.Count > 0 ?
                            (from p in file.Properties
                             select new V201508.StringDictionaryItem
                             {
                                 Key = p.Key,
                                 Value = p.Value
                             }).ToArray() : null,
                         Security = file.Security.FromTemplateToSchemaObjectSecurityV201508()
                     }).ToArray();
            }
            else
            {
                result.Files = null;
            }

            #endregion

            #region Pages

            // Translate Pages, if any
            if (template.Pages != null && template.Pages.Count > 0)
            {
                var pages = new List<V201508.Page>();

                foreach (var page in template.Pages)
                {
                    var schemaPage = new V201508.Page();

                    var pageLayout = V201508.WikiPageLayout.OneColumn;
                    switch (page.Layout)
                    {
                        case WikiPageLayout.OneColumn:
                            pageLayout = V201508.WikiPageLayout.OneColumn;
                            break;
                        case WikiPageLayout.OneColumnSideBar:
                            pageLayout = V201508.WikiPageLayout.OneColumnSidebar;
                            break;
                        case WikiPageLayout.TwoColumns:
                            pageLayout = V201508.WikiPageLayout.TwoColumns;
                            break;
                        case WikiPageLayout.TwoColumnsHeader:
                            pageLayout = V201508.WikiPageLayout.TwoColumnsHeader;
                            break;
                        case WikiPageLayout.TwoColumnsHeaderFooter:
                            pageLayout = V201508.WikiPageLayout.TwoColumnsHeaderFooter;
                            break;
                        case WikiPageLayout.ThreeColumns:
                            pageLayout = V201508.WikiPageLayout.ThreeColumns;
                            break;
                        case WikiPageLayout.ThreeColumnsHeader:
                            pageLayout = V201508.WikiPageLayout.ThreeColumnsHeader;
                            break;
                        case WikiPageLayout.ThreeColumnsHeaderFooter:
                            pageLayout = V201508.WikiPageLayout.ThreeColumnsHeaderFooter;
                            break;
                    }
                    schemaPage.Layout = pageLayout;
                    schemaPage.Overwrite = page.Overwrite;
                    schemaPage.Security = (page.Security != null) ? page.Security.FromTemplateToSchemaObjectSecurityV201508() : null;

                    schemaPage.WebParts = page.WebParts.Count > 0 ?
                        (from wp in page.WebParts
                         select new V201508.WikiPageWebPart
                         {
                             Column = (int)wp.Column,
                             Row = (int)wp.Row,
                             Contents = wp.Contents,
                             Title = wp.Title,
                         }).ToArray() : null;

                    schemaPage.Url = page.Url;

                    pages.Add(schemaPage);
                }

                result.Pages = new V201508.Pages()
                {
                    Page = pages.ToArray(),
#pragma warning disable 618
                    WelcomePage = template.Pages.Any(p => p.WelcomePage = true) ? template.Pages.Last(p => p.WelcomePage = true).Url : null
#pragma warning restore 618
                };
            }

            #endregion

            #region Taxonomy

            // Translate Taxonomy elements, if any
            if (template.TermGroups != null && template.TermGroups.Count > 0)
            {
                result.TermGroups =
                    (from grp in template.TermGroups
                     select new V201508.TermGroup
                     {
                         Name = grp.Name,
                         ID = grp.Id.ToString(),
                         Description = grp.Description,
                         TermSets = (
                            from termSet in grp.TermSets
                            select new V201508.TermSet
                            {
                                ID = termSet.Id.ToString(),
                                Name = termSet.Name,
                                IsAvailableForTagging = termSet.IsAvailableForTagging,
                                IsOpenForTermCreation = termSet.IsOpenForTermCreation,
                                Description = termSet.Description,
                                Language = termSet.Language.HasValue ? termSet.Language.Value : 0,
                                LanguageSpecified = termSet.Language.HasValue,
                                Terms = termSet.Terms.FromModelTermsToSchemaTermsV201508(),
                                CustomProperties = termSet.Properties.Count > 0 ?
                                     (from p in termSet.Properties
                                      select new V201508.StringDictionaryItem
                                      {
                                          Key = p.Key,
                                          Value = p.Value
                                      }).ToArray() : null,
                            }).ToArray(),
                     }).ToArray();
            }

            #endregion

            #region Composed Looks

            // Translate ComposedLook, if any
            if (template.ComposedLook != null && !template.ComposedLook.Equals(Model.ComposedLook.Empty))
            {
                result.ComposedLook = new V201508.ComposedLook
                {
#pragma warning disable 618
                    AlternateCSS = template.ComposedLook.AlternateCSS,
#pragma warning restore 618
                    BackgroundFile = template.ComposedLook.BackgroundFile,
                    ColorFile = template.ComposedLook.ColorFile,
                    FontFile = template.ComposedLook.FontFile,
#pragma warning disable 618
                    MasterPage = template.ComposedLook.MasterPage,
#pragma warning restore 618
                    Name = template.ComposedLook.Name,
#pragma warning disable 618
                    SiteLogo = template.ComposedLook.SiteLogo,
#pragma warning restore 618
                    Version = template.ComposedLook.Version,
                    VersionSpecified = true,
                };
            }

            #endregion

            #region Workflows

            if (template.Workflows != null &&
                (template.Workflows.WorkflowDefinitions.Any() || template.Workflows.WorkflowSubscriptions.Any()))
            {
                result.Workflows = new V201508.Workflows
                {
                    WorkflowDefinitions = template.Workflows.WorkflowDefinitions.Count > 0 ?
                        (from wd in template.Workflows.WorkflowDefinitions
                         select new WorkflowsWorkflowDefinition
                         {
                             AssociationUrl = wd.AssociationUrl,
                             Description = wd.Description,
                             DisplayName = wd.DisplayName,
                             DraftVersion = wd.DraftVersion,
                             FormField = (wd.FormField != null) ? wd.FormField.ToXmlElement() : null,
                             Id = wd.Id.ToString(),
                             InitiationUrl = wd.InitiationUrl,
                             Properties = (wd.Properties != null && wd.Properties.Count > 0) ?
                                (from p in wd.Properties
                                 select new V201508.StringDictionaryItem
                                 {
                                     Key = p.Key,
                                     Value = p.Value,
                                 }).ToArray() : null,
                             Published = wd.Published,
                             PublishedSpecified = true,
                             RequiresAssociationForm = wd.RequiresAssociationForm,
                             RequiresAssociationFormSpecified = true,
                             RequiresInitiationForm = wd.RequiresInitiationForm,
                             RequiresInitiationFormSpecified = true,
                             RestrictToScope = wd.RestrictToScope,
                             RestrictToType = (V201508.WorkflowsWorkflowDefinitionRestrictToType)Enum.Parse(typeof(V201508.WorkflowsWorkflowDefinitionRestrictToType), wd.RestrictToType),
                             RestrictToTypeSpecified = true,
                             XamlPath = wd.XamlPath,
                         }).ToArray() : null,
                    WorkflowSubscriptions = template.Workflows.WorkflowSubscriptions.Count > 0 ?
                        (from ws in template.Workflows.WorkflowSubscriptions
                         select new WorkflowsWorkflowSubscription
                         {
                             DefinitionId = ws.DefinitionId.ToString(),
                             Enabled = ws.Enabled,
                             EventSourceId = ws.EventSourceId,
                             ItemAddedEvent = ws.EventTypes.Contains("ItemAdded"),
                             ItemUpdatedEvent = ws.EventTypes.Contains("ItemUpdated"),
                             WorkflowStartEvent = ws.EventTypes.Contains("WorkflowStart"),
                             ListId = ws.ListId,
                             ManualStartBypassesActivationLimit = ws.ManualStartBypassesActivationLimit,
                             ManualStartBypassesActivationLimitSpecified = true,
                             Name = ws.Name,
                             ParentContentTypeId = ws.ParentContentTypeId,
                             PropertyDefinitions = (ws.PropertyDefinitions != null && ws.PropertyDefinitions.Count > 0) ?
                                (from pd in ws.PropertyDefinitions
                                 select new V201508.StringDictionaryItem
                                 {
                                     Key = pd.Key,
                                     Value = pd.Value,
                                 }).ToArray() : null,
                             StatusFieldName = ws.StatusFieldName,
                         }).ToArray() : null,
                };
            }
            else
            {
                result.Workflows = null;
            }

            #endregion

            #region Search Settings

#pragma warning disable 618
            if (!String.IsNullOrEmpty(template.SearchSettings))
            {
                result.SearchSettings = template.SearchSettings.ToXmlElement();
            }
#pragma warning restore 618
            #endregion

            #region Publishing

            if (template.Publishing != null)
            {
                result.Publishing = new V201508.Publishing
                {
                    AutoCheckRequirements = (V201508.PublishingAutoCheckRequirements)Enum.Parse(typeof(V201508.PublishingAutoCheckRequirements), template.Publishing.AutoCheckRequirements.ToString()),
                    AvailableWebTemplates = template.Publishing.AvailableWebTemplates.Count > 0 ?
                        (from awt in template.Publishing.AvailableWebTemplates
                         select new V201508.PublishingWebTemplate
                         {
                             LanguageCode = awt.LanguageCode,
                             LanguageCodeSpecified = true,
                             TemplateName = awt.TemplateName,
                         }).ToArray() : null,
                    DesignPackage = template.Publishing.DesignPackage != null ? new V201508.PublishingDesignPackage
                    {
                        DesignPackagePath = template.Publishing.DesignPackage.DesignPackagePath,
                        MajorVersion = template.Publishing.DesignPackage.MajorVersion,
                        MajorVersionSpecified = true,
                        MinorVersion = template.Publishing.DesignPackage.MinorVersion,
                        MinorVersionSpecified = true,
                        PackageGuid = template.Publishing.DesignPackage.PackageGuid.ToString(),
                        PackageName = template.Publishing.DesignPackage.PackageName,
                    } : null,
                    PageLayouts = template.Publishing.PageLayouts != null ?
                        new V201508.PublishingPageLayouts
                        {
                            PageLayout = template.Publishing.PageLayouts.Count > 0 ?
                        (from pl in template.Publishing.PageLayouts
                         select new V201508.PublishingPageLayoutsPageLayout
                         {
                             Path = pl.Path,
                         }).ToArray() : null,
                            Default = template.Publishing.PageLayouts.Any(p => p.IsDefault) ?
                                template.Publishing.PageLayouts.Last(p => p.IsDefault).Path : null,
                        } : null,
                };
            }
            else
            {
                result.Publishing = null;
            }

            #endregion

            #region AddIns

            if (template.AddIns != null && template.AddIns.Count > 0)
            {
                result.AddIns =
                    (from addin in template.AddIns
                     select new V201508.AddInsAddin
                     {
                         PackagePath = addin.PackagePath,
                         Source = (V201508.AddInsAddinSource)Enum.Parse(typeof(V201508.AddInsAddinSource), addin.Source),
                     }).ToArray();
            }
            else
            {
                result.AddIns = null;
            }

            #endregion

            #region Providers

            // Translate Providers, if any
#pragma warning disable 618
            if ((template.Providers != null && template.Providers.Count > 0) || (template.ExtensibilityHandlers != null && template.ExtensibilityHandlers.Count > 0))
             {
                var extensibilityHandlers = template.ExtensibilityHandlers.Union(template.Providers);
                result.Providers =
                    (from provider in extensibilityHandlers
                     select new V201508.Provider
                     {
                         HandlerType = String.Format("{0}, {1}", provider.Type, provider.Assembly),
                         Configuration = provider.Configuration != null ? provider.Configuration.ToXmlNode() : null,
                         Enabled = provider.Enabled,
                     }).ToArray();
            }
            else
            {
                result.Providers = null;
            }
#pragma warning restore 618
            #endregion

            XmlSerializerNamespaces ns =
                new XmlSerializerNamespaces();
            ns.Add(((IXMLSchemaFormatter)this).NamespacePrefix,
                ((IXMLSchemaFormatter)this).NamespaceUri);

            var output = XMLSerializer.SerializeToStream<V201508.Provisioning>(wrappedResult, ns);
            output.Position = 0;
            return (output);
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(Stream template, String identifier)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            // Crate a copy of the source stream
            MemoryStream sourceStream = new MemoryStream();
            template.CopyTo(sourceStream);
            sourceStream.Position = 0;

            // Check the provided template against the XML schema
            if (!this.IsValid(sourceStream))
            {
                // TODO: Use resource file
                throw new ApplicationException("The provided template is not valid!");
            }

            sourceStream.Position = 0;
            XDocument xml = XDocument.Load(sourceStream);
#pragma warning disable 618
            XNamespace pnp = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_08;
#pragma warning restore 618

            // Prepare a variable to hold the single source formatted template
            V201508.ProvisioningTemplate source = null;

            // Prepare a variable to hold the resulting ProvisioningTemplate instance
            Model.ProvisioningTemplate result = new Model.ProvisioningTemplate();

            // Determine if we're working on a wrapped SharePointProvisioningTemplate or not
            if (xml.Root.Name == pnp + "Provisioning")
            {
                // Deserialize the whole wrapper
                V201508.Provisioning wrappedResult = XMLSerializer.Deserialize<V201508.Provisioning>(xml);

                // Handle the wrapper schema parameters
                if (wrappedResult.Preferences != null &&
                    wrappedResult.Preferences.Parameters != null &&
                    wrappedResult.Preferences.Parameters.Length > 0)
                {
                    foreach (var parameter in wrappedResult.Preferences.Parameters)
                    {
                        result.Parameters.Add(parameter.Key, parameter.Text != null ? parameter.Text.Aggregate(String.Empty, (acc, i) => acc + i) : null);
                    }
                }

                foreach (var templates in wrappedResult.Templates)
                {
                    // Let's see if we have an in-place template with the provided ID or if we don't have a provided ID at all
                    source = templates.ProvisioningTemplate.FirstOrDefault(spt => spt.ID == identifier || String.IsNullOrEmpty(identifier));

                    // If we don't have a template, but there are external file references
                    if (source == null && templates.ProvisioningTemplateFile.Length > 0)
                    {
                        // Otherwise let's see if we have an external file for the template
                        var externalSource = templates.ProvisioningTemplateFile.FirstOrDefault(sptf => sptf.ID == identifier);

                        Stream externalFileStream = this._provider.Connector.GetFileStream(externalSource.File);
                        xml = XDocument.Load(externalFileStream);

                        if (xml.Root.Name != pnp + "ProvisioningTemplate")
                        {
                            throw new ApplicationException("Invalid external file format. Expected a ProvisioningTemplate file!");
                        }
                        else
                        {
                            source = XMLSerializer.Deserialize<V201508.ProvisioningTemplate>(xml);
                        }
                    }

                    if (source != null)
                    {
                        break;
                    }
                }
            }
            else if (xml.Root.Name == pnp + "ProvisioningTemplate")
            {
                var IdAttribute = xml.Root.Attribute("ID");

                // If there is a provided ID, and if it doesn't equal the current ID
                if (!String.IsNullOrEmpty(identifier) &&
                    IdAttribute != null &&
                    IdAttribute.Value != identifier)
                {
                    // TODO: Use resource file
                    throw new ApplicationException("The provided template identifier is not available!");
                }
                else
                {
                    source = XMLSerializer.Deserialize<V201508.ProvisioningTemplate>(xml);
                }
            }

            #region Basic Properties

            // Translate basic properties
            result.Id = source.ID;
            result.Version = (Double)source.Version;
            result.SitePolicy = source.SitePolicy;
            result.ImagePreviewUrl = source.ImagePreviewUrl;
            result.DisplayName = source.DisplayName;
            result.Description = source.Description;

            if (source.Properties != null && source.Properties.Length > 0)
            {
                result.Properties.AddRange(
                    (from p in source.Properties
                     select p).ToDictionary(i => i.Key, i => i.Value));
            }

            #endregion

            #region Property Bag

            // Translate PropertyBagEntries, if any
            if (source.PropertyBagEntries != null)
            {
                result.PropertyBagEntries.AddRange(
                    from bag in source.PropertyBagEntries
                    select new Model.PropertyBagEntry
                    {
                        Key = bag.Key,
                        Value = bag.Value,
                        Indexed = bag.Indexed,
                        Overwrite = bag.OverwriteSpecified ? bag.Overwrite : false,
                    });
            }

            #endregion

            #region Regional Settings

            if (source.RegionalSettings != null)
            {
                result.RegionalSettings = new Model.RegionalSettings()
                {
                    AdjustHijriDays = source.RegionalSettings.AdjustHijriDaysSpecified ? source.RegionalSettings.AdjustHijriDays : 0,
                    AlternateCalendarType = source.RegionalSettings.AlternateCalendarTypeSpecified ? source.RegionalSettings.AlternateCalendarType.FromSchemaToTemplateCalendarTypeV201508() : Microsoft.SharePoint.Client.CalendarType.None,
                    CalendarType = source.RegionalSettings.CalendarTypeSpecified ? source.RegionalSettings.CalendarType.FromSchemaToTemplateCalendarTypeV201508() : Microsoft.SharePoint.Client.CalendarType.None,
                    Collation = source.RegionalSettings.CollationSpecified ? source.RegionalSettings.Collation : 0,
                    FirstDayOfWeek = source.RegionalSettings.FirstDayOfWeekSpecified ?
                        (System.DayOfWeek)Enum.Parse(typeof(System.DayOfWeek), source.RegionalSettings.FirstDayOfWeek.ToString()) : System.DayOfWeek.Sunday,
                    FirstWeekOfYear = source.RegionalSettings.FirstWeekOfYearSpecified ? source.RegionalSettings.FirstWeekOfYear : 0,
                    LocaleId = source.RegionalSettings.LocaleIdSpecified ? source.RegionalSettings.LocaleId : 1033,
                    ShowWeeks = source.RegionalSettings.ShowWeeksSpecified ? source.RegionalSettings.ShowWeeks : false,
                    Time24 = source.RegionalSettings.Time24Specified ? source.RegionalSettings.Time24 : false,
                    TimeZone = Int32.Parse(source.RegionalSettings.TimeZone),
                    WorkDayEndHour = source.RegionalSettings.WorkDayEndHourSpecified ? source.RegionalSettings.WorkDayEndHour.FromSchemaToTemplateWorkHourV201508() : Model.WorkHour.PM0600,
                    WorkDays = source.RegionalSettings.WorkDaysSpecified ? source.RegionalSettings.WorkDays : 5,
                    WorkDayStartHour = source.RegionalSettings.WorkDayStartHourSpecified ? source.RegionalSettings.WorkDayStartHour.FromSchemaToTemplateWorkHourV201508() : Model.WorkHour.AM0900,
                };
            }
            else
            {
                result.RegionalSettings = null;
            }

            #endregion

            #region Supported UI Languages

            if (source.SupportedUILanguages != null && source.SupportedUILanguages.Length > 0)
            {
                result.SupportedUILanguages.AddRange(
                    from l in source.SupportedUILanguages
                    select new SupportedUILanguage
                    {
                        LCID = l.LCID,
                    });
            }

            #endregion

            #region Audit Settings

            if (source.AuditSettings != null)
            {
                result.AuditSettings = new Model.AuditSettings
                {
                    AuditLogTrimmingRetention = source.AuditSettings.AuditLogTrimmingRetentionSpecified ? source.AuditSettings.AuditLogTrimmingRetention : 0,
                    TrimAuditLog = source.AuditSettings.TrimAuditLogSpecified ? source.AuditSettings.TrimAuditLog : false,
                    AuditFlags = source.AuditSettings.Audit.Aggregate(Microsoft.SharePoint.Client.AuditMaskType.None, (acc, next) => acc |= (Microsoft.SharePoint.Client.AuditMaskType)Enum.Parse(typeof(Microsoft.SharePoint.Client.AuditMaskType), next.AuditFlag.ToString())),
                };
            }

            #endregion

            #region Security

            // Translate Security configuration, if any
            if (source.Security != null)
            {
                if (source.Security.AdditionalAdministrators != null)
                {
                    result.Security.AdditionalAdministrators.AddRange(
                        from user in source.Security.AdditionalAdministrators
                        select new Model.User
                        {
                            Name = user.Name,
                        });
                }
                if (source.Security.AdditionalOwners != null)
                {
                    result.Security.AdditionalOwners.AddRange(
                        from user in source.Security.AdditionalOwners
                        select new Model.User
                        {
                            Name = user.Name,
                        });
                }
                if (source.Security.AdditionalMembers != null)
                {
                    result.Security.AdditionalMembers.AddRange(
                        from user in source.Security.AdditionalMembers
                        select new Model.User
                        {
                            Name = user.Name,
                        });
                }
                if (source.Security.AdditionalVisitors != null)
                {
                    result.Security.AdditionalVisitors.AddRange(
                        from user in source.Security.AdditionalVisitors
                        select new Model.User
                        {
                            Name = user.Name,
                        });
                }
                if (source.Security.SiteGroups != null)
                {
                    result.Security.SiteGroups.AddRange(
                        from g in source.Security.SiteGroups
                        select new Model.SiteGroup(g.Members != null ? from m in g.Members select new Model.User { Name = m.Name } : null)
                        {
                            AllowMembersEditMembership = g.AllowMembersEditMembershipSpecified ? g.AllowMembersEditMembership : false,
                            AllowRequestToJoinLeave = g.AllowRequestToJoinLeaveSpecified ? g.AllowRequestToJoinLeave : false,
                            AutoAcceptRequestToJoinLeave = g.AutoAcceptRequestToJoinLeaveSpecified ? g.AutoAcceptRequestToJoinLeave : false,
                            Description = g.Description,
                            OnlyAllowMembersViewMembership = g.OnlyAllowMembersViewMembershipSpecified ? g.OnlyAllowMembersViewMembership : false,
                            Owner = g.Owner,
                            RequestToJoinLeaveEmailSetting = g.RequestToJoinLeaveEmailSetting,
                            Title = g.Title,
                        });
                }
                if (source.Security.Permissions != null)
                {
                    if (source.Security.Permissions.RoleAssignments != null && source.Security.Permissions.RoleAssignments.Length > 0)
                    {
                        result.Security.SiteSecurityPermissions.RoleAssignments.AddRange
                            (from ra in source.Security.Permissions.RoleAssignments
                             select new Model.RoleAssignment
                             {
                                 Principal = ra.Principal,
                                 RoleDefinition = ra.RoleDefinition,
                             });
                    }
                    if (source.Security.Permissions.RoleDefinitions != null && source.Security.Permissions.RoleDefinitions.Length > 0)
                    {
                        result.Security.SiteSecurityPermissions.RoleDefinitions.AddRange
                            (from rd in source.Security.Permissions.RoleDefinitions
                             select new Model.RoleDefinition(
                                 from p in rd.Permissions
                                 select (Microsoft.SharePoint.Client.PermissionKind)Enum.Parse(typeof(Microsoft.SharePoint.Client.PermissionKind), p.ToString()))
                             {
                                 Description = rd.Description,
                                 Name = rd.Name,
                             });
                    }
                }
            }

            #endregion

            #region Site Columns

            // Translate Site Columns (Fields), if any
            if ((source.SiteFields != null) && (source.SiteFields.Any != null))
            {
                result.SiteFields.AddRange(
                    from field in source.SiteFields.Any
                    select new Field
                    {
                        SchemaXml = field.OuterXml,
                    });
            }

            #endregion

            #region Content Types

            // Translate ContentTypes, if any
            if ((source.ContentTypes != null) && (source.ContentTypes != null))
            {
                result.ContentTypes.AddRange(
                    from contentType in source.ContentTypes
                    select new ContentType(
                        contentType.ID,
                        contentType.Name,
                        contentType.Description,
                        contentType.Group,
                        contentType.Sealed,
                        contentType.Hidden,
                        contentType.ReadOnly,
                        (contentType.DocumentTemplate != null ?
                            contentType.DocumentTemplate.TargetName : null),
                        contentType.Overwrite,
                        (contentType.FieldRefs != null ?
                            (from fieldRef in contentType.FieldRefs
                             select new Model.FieldRef(fieldRef.Name)
                             {
                                 Id = Guid.Parse(fieldRef.ID),
                                 Hidden = fieldRef.Hidden,
                                 Required = fieldRef.Required
                             }) : null)
                        )
                    {
                        DocumentSetTemplate = contentType.DocumentSetTemplate != null ?
                            new Model.DocumentSetTemplate(
                                contentType.DocumentSetTemplate.WelcomePage,
                                contentType.DocumentSetTemplate.AllowedContentTypes != null ?
                                    (from act in contentType.DocumentSetTemplate.AllowedContentTypes
                                     select act.ContentTypeID) : null,
                                contentType.DocumentSetTemplate.DefaultDocuments != null ?
                                    (from dd in contentType.DocumentSetTemplate.DefaultDocuments
                                     select new Model.DefaultDocument
                                     {
                                         ContentTypeId = dd.ContentTypeID,
                                         FileSourcePath = dd.FileSourcePath,
                                         Name = dd.Name,
                                     }) : null,
                                contentType.DocumentSetTemplate.SharedFields != null ?
                                    (from sf in contentType.DocumentSetTemplate.SharedFields
                                     select Guid.Parse(sf.ID)) : null,
                                contentType.DocumentSetTemplate.WelcomePageFields != null ?
                                    (from wpf in contentType.DocumentSetTemplate.WelcomePageFields
                                     select Guid.Parse(wpf.ID)) : null
                                ) : null,
                        DisplayFormUrl = contentType.DisplayFormUrl,
                        EditFormUrl = contentType.EditFormUrl,
                        NewFormUrl = contentType.NewFormUrl,
                    }
                );
            }

            #endregion

            #region List Instances

            // Translate Lists Instances, if any
            if (source.Lists != null)
            {
                result.Lists.AddRange(
                    from list in source.Lists
                    select new Model.ListInstance(
                        (list.ContentTypeBindings != null ?
                                (from contentTypeBinding in list.ContentTypeBindings
                                 select new Model.ContentTypeBinding
                                 {
                                     ContentTypeId = contentTypeBinding.ContentTypeID,
                                     Default = contentTypeBinding.Default,
                                 }) : null),
                        (list.Views != null ?
                                (from view in list.Views.Any
                                 select new View
                                 {
                                     SchemaXml = view.OuterXml,
                                 }) : null),
                        (list.Fields != null ?
                                (from field in list.Fields.Any
                                 select new Field
                                 {
                                     SchemaXml = field.OuterXml,
                                 }) : null),
                        (list.FieldRefs != null ?
                                    (from fieldRef in list.FieldRefs
                                     select new Model.FieldRef(fieldRef.Name)
                                     {
                                         DisplayName = fieldRef.DisplayName,
                                         Hidden = fieldRef.Hidden,
                                         Required = fieldRef.Required,
                                         Id = Guid.Parse(fieldRef.ID)
                                     }) : null),
                        (list.DataRows != null ?
                                    (from dataRow in list.DataRows
                                     select new Model.DataRow(
                                 (from dataValue in dataRow.DataValue
                                  select dataValue).ToDictionary(k => k.FieldName, v => v.Value),
                                 dataRow.Security.FromSchemaToTemplateObjectSecurityV201508()
                             )).ToList() : null),
                        (list.FieldDefaults != null ?
                            (from fd in list.FieldDefaults
                             select fd).ToDictionary(k => k.FieldName, v => v.Value) : null),
                        list.Security.FromSchemaToTemplateObjectSecurityV201508()
                        )
                    {
                        ContentTypesEnabled = list.ContentTypesEnabled,
                        Description = list.Description,
                        DocumentTemplate = list.DocumentTemplate,
                        EnableVersioning = list.EnableVersioning,
                        EnableMinorVersions = list.EnableMinorVersions,
                        DraftVersionVisibility = list.DraftVersionVisibility,
                        EnableModeration = list.EnableModeration,
                        Hidden = list.Hidden,
                        MinorVersionLimit = list.MinorVersionLimitSpecified ? list.MinorVersionLimit : 0,
                        MaxVersionLimit = list.MaxVersionLimitSpecified ? list.MaxVersionLimit : 0,
                        OnQuickLaunch = list.OnQuickLaunch,
                        EnableAttachments = list.EnableAttachments,
                        EnableFolderCreation = list.EnableFolderCreation,
                        RemoveExistingContentTypes = list.RemoveExistingContentTypes,
                        TemplateFeatureID = !String.IsNullOrEmpty(list.TemplateFeatureID) ? Guid.Parse(list.TemplateFeatureID) : Guid.Empty,
                        RemoveExistingViews = list.Views != null ? list.Views.RemoveExistingViews : false,
                        TemplateType = list.TemplateType,
                        Title = list.Title,
                        Url = list.Url,
                    });
            }

            #endregion

            #region Features

            // Translate Features, if any
            if (source.Features != null)
            {
                if (result.Features.SiteFeatures != null && source.Features.SiteFeatures != null)
                {
                    result.Features.SiteFeatures.AddRange(
                        from feature in source.Features.SiteFeatures
                        select new Model.Feature
                        {
                            Id = new Guid(feature.ID),
                            Deactivate = feature.Deactivate,
                        });
                }
                if (result.Features.WebFeatures != null && source.Features.WebFeatures != null)
                {
                    result.Features.WebFeatures.AddRange(
                        from feature in source.Features.WebFeatures
                        select new Model.Feature
                        {
                            Id = new Guid(feature.ID),
                            Deactivate = feature.Deactivate,
                        });
                }
            }

            #endregion

            #region Custom Actions

            // Translate CustomActions, if any
            if (source.CustomActions != null)
            {
                if (result.CustomActions.SiteCustomActions != null && source.CustomActions.SiteCustomActions != null)
                {
                    result.CustomActions.SiteCustomActions.AddRange(
                        from customAction in source.CustomActions.SiteCustomActions
                        select new Model.CustomAction
                        {
                            CommandUIExtension = (customAction.CommandUIExtension != null && customAction.CommandUIExtension.Any != null) ?
                                (new XElement("CommandUIExtension", from x in customAction.CommandUIExtension.Any select x.ToXElement())) : null,
                            Description = customAction.Description,
                            Enabled = customAction.Enabled,
                            Group = customAction.Group,
                            ImageUrl = customAction.ImageUrl,
                            Location = customAction.Location,
                            Name = customAction.Name,
#pragma warning disable 618
                            RightsValue = customAction.RightsSpecified ? customAction.Rights : 0,
#pragma warning restore 618
                            ScriptBlock = customAction.ScriptBlock,
                            ScriptSrc = customAction.ScriptSrc,
                            Sequence = customAction.SequenceSpecified ? customAction.Sequence : 100,
                            Title = customAction.Title,
                            Url = customAction.Url,
                        });
                }
                if (result.CustomActions.WebCustomActions != null && source.CustomActions.WebCustomActions != null)
                {
                    result.CustomActions.WebCustomActions.AddRange(
                        from customAction in source.CustomActions.WebCustomActions
                        select new Model.CustomAction
                        {
                            CommandUIExtension = (customAction.CommandUIExtension != null && customAction.CommandUIExtension.Any != null) ?
                                (new XElement("CommandUIExtension", from x in customAction.CommandUIExtension.Any select x.ToXElement())) : null,
                            Description = customAction.Description,
                            Enabled = customAction.Enabled,
                            Group = customAction.Group,
                            ImageUrl = customAction.ImageUrl,
                            Location = customAction.Location,
                            Name = customAction.Name,
#pragma warning disable 618
                            RightsValue = customAction.RightsSpecified ? customAction.Rights : 0,
#pragma warning restore 618
                            ScriptBlock = customAction.ScriptBlock,
                            ScriptSrc = customAction.ScriptSrc,
                            Sequence = customAction.SequenceSpecified ? customAction.Sequence : 100,
                            Title = customAction.Title,
                            Url = customAction.Url,
                        });
                }
            }

            #endregion

            #region Files

            // Translate Files, if any
            if (source.Files != null)
            {
                result.Files.AddRange(
                    from file in source.Files
                    select new Model.File(file.Src,
                        file.Folder,
                        file.Overwrite,
                        file.WebParts != null ?
                            (from wp in file.WebParts
                             select new Model.WebPart
                             {
                                 Order = (uint)wp.Order,
                                 Zone = wp.Zone,
                                 Title = wp.Title,
                                 Contents = wp.Contents
                             }) : null,
                        file.Properties != null ? file.Properties.ToDictionary(k => k.Key, v => v.Value) : null,
                        file.Security.FromSchemaToTemplateObjectSecurityV201508()
                        )
                    );
            }

            #endregion

            #region Pages

            // Translate Pages, if any
            if (source.Pages != null)
            {
                foreach (var page in source.Pages.Page)
                {

                    var pageLayout = WikiPageLayout.OneColumn;
                    switch (page.Layout)
                    {
                        case V201508.WikiPageLayout.OneColumn:
                            pageLayout = WikiPageLayout.OneColumn;
                            break;
                        case V201508.WikiPageLayout.OneColumnSidebar:
                            pageLayout = WikiPageLayout.OneColumnSideBar;
                            break;
                        case V201508.WikiPageLayout.TwoColumns:
                            pageLayout = WikiPageLayout.TwoColumns;
                            break;
                        case V201508.WikiPageLayout.TwoColumnsHeader:
                            pageLayout = WikiPageLayout.TwoColumnsHeader;
                            break;
                        case V201508.WikiPageLayout.TwoColumnsHeaderFooter:
                            pageLayout = WikiPageLayout.TwoColumnsHeaderFooter;
                            break;
                        case V201508.WikiPageLayout.ThreeColumns:
                            pageLayout = WikiPageLayout.ThreeColumns;
                            break;
                        case V201508.WikiPageLayout.ThreeColumnsHeader:
                            pageLayout = WikiPageLayout.ThreeColumnsHeader;
                            break;
                        case V201508.WikiPageLayout.ThreeColumnsHeaderFooter:
                            pageLayout = WikiPageLayout.ThreeColumnsHeaderFooter;
                            break;
                    }

#pragma warning disable 618
                    result.Pages.Add(new Model.Page(page.Url, page.Overwrite, pageLayout,
                        (page.WebParts != null ?
                            (from wp in page.WebParts
                             select new Model.WebPart
                             {
                                 Title = wp.Title,
                                 Column = (uint)wp.Column,
                                 Row = (uint)wp.Row,
                                 Contents = wp.Contents

                             }).ToList() : null),
                        source.Pages.WelcomePage == page.Url,
                        page.Security.FromSchemaToTemplateObjectSecurityV201508()));
                }
#pragma warning restore 618

            }


            #endregion

            #region Taxonomy

            // Translate Termgroups, if any
            if (source.TermGroups != null)
            {
                result.TermGroups.AddRange(
                    from termGroup in source.TermGroups
                    select new Model.TermGroup(
                        !string.IsNullOrEmpty(termGroup.ID) ? Guid.Parse(termGroup.ID) : Guid.Empty,
                        termGroup.Name,
                        new List<Model.TermSet>(
                            from termSet in termGroup.TermSets
                            select new Model.TermSet(
                                !string.IsNullOrEmpty(termSet.ID) ? Guid.Parse(termSet.ID) : Guid.Empty,
                                termSet.Name,
                                termSet.LanguageSpecified ? (int?)termSet.Language : null,
                                termSet.IsAvailableForTagging,
                                termSet.IsOpenForTermCreation,
                                termSet.Terms != null ? termSet.Terms.FromSchemaTermsToModelTermsV201508() : null,
                                termSet.CustomProperties != null ? termSet.CustomProperties.ToDictionary(k => k.Key, v => v.Value) : null)
                            {
                                Description = termSet.Description,
                            })
                        )
                    {
                        Description = termGroup.Description,
                    });
            }

            #endregion

            #region Composed Looks

            // Translate ComposedLook, if any
            if (source.ComposedLook != null)
            {
#pragma warning disable 618
                result.ComposedLook.AlternateCSS = source.ComposedLook.AlternateCSS;
#pragma warning restore 618
                result.ComposedLook.BackgroundFile = source.ComposedLook.BackgroundFile;
                result.ComposedLook.ColorFile = source.ComposedLook.ColorFile;
                result.ComposedLook.FontFile = source.ComposedLook.FontFile;
#pragma warning disable 618
                result.ComposedLook.MasterPage = source.ComposedLook.MasterPage;
#pragma warning restore 618
                result.ComposedLook.Name = source.ComposedLook.Name;
#pragma warning disable 618
                result.ComposedLook.SiteLogo = source.ComposedLook.SiteLogo;
#pragma warning restore 618
                result.ComposedLook.Version = source.ComposedLook.Version;
            }

            #endregion

            #region Workflows

            if (source.Workflows != null)
            {
                result.Workflows = new Model.Workflows(
                    source.Workflows.WorkflowDefinitions.Length > 0 ?
                        (from wd in source.Workflows.WorkflowDefinitions
                         select new Model.WorkflowDefinition(
                             (wd.Properties != null && wd.Properties.Length > 0) ?
                             (from p in wd.Properties
                              select p).ToDictionary(i => i.Key, i => i.Value) : null)
                         {
                             AssociationUrl = wd.AssociationUrl,
                             Description = wd.Description,
                             DisplayName = wd.DisplayName,
                             DraftVersion = wd.DraftVersion,
                             FormField = wd.FormField.OuterXml,
                             Id = Guid.Parse(wd.Id),
                             InitiationUrl = wd.InitiationUrl,
                             Published = wd.PublishedSpecified ? wd.Published : false,
                             RequiresAssociationForm = wd.RequiresAssociationFormSpecified ? wd.RequiresAssociationForm : false,
                             RequiresInitiationForm = wd.RequiresInitiationFormSpecified ? wd.RequiresInitiationForm : false,
                             RestrictToScope = wd.RestrictToScope,
                             RestrictToType = wd.RestrictToType.ToString(),
                             XamlPath = wd.XamlPath,
                         }) : null,
                    source.Workflows.WorkflowSubscriptions.Length > 0 ?
                        (from ws in source.Workflows.WorkflowSubscriptions
                         select new Model.WorkflowSubscription(
                             (ws.PropertyDefinitions != null && ws.PropertyDefinitions.Length > 0) ?
                             (from pd in ws.PropertyDefinitions
                              select pd).ToDictionary(i => i.Key, i => i.Value) : null)
                         {
                             DefinitionId = Guid.Parse(ws.DefinitionId),
                             Enabled = ws.Enabled,
                             EventSourceId = ws.EventSourceId,
                             EventTypes = (new String[] {
                                ws.ItemAddedEvent? "ItemAdded" : null,
                                ws.ItemUpdatedEvent? "ItemUpdated" : null,
                                ws.WorkflowStartEvent? "WorkflowStart" : null }).Where(e => e != null).ToList(),
                             ListId = ws.ListId,
                             ManualStartBypassesActivationLimit = ws.ManualStartBypassesActivationLimitSpecified ? ws.ManualStartBypassesActivationLimit : false,
                             Name = ws.Name,
                             ParentContentTypeId = ws.ParentContentTypeId,
                             StatusFieldName = ws.StatusFieldName,
                         }) : null
                    );
            }

            #endregion

            #region Search Settings

            if (source.SearchSettings != null)
            {
#pragma warning disable 618
                result.SearchSettings = source.SearchSettings.OuterXml;
#pragma warning restore 618
            }

            #endregion

            #region Publishing

            if (source.Publishing != null)
            {
                result.Publishing = new Model.Publishing(
                    (Model.AutoCheckRequirementsOptions)Enum.Parse(typeof(Model.AutoCheckRequirementsOptions), source.Publishing.AutoCheckRequirements.ToString()),
                    source.Publishing.DesignPackage != null ?
                    new Model.DesignPackage
                    {
                        DesignPackagePath = source.Publishing.DesignPackage.DesignPackagePath,
                        MajorVersion = source.Publishing.DesignPackage.MajorVersionSpecified ? source.Publishing.DesignPackage.MajorVersion : 0,
                        MinorVersion = source.Publishing.DesignPackage.MinorVersionSpecified ? source.Publishing.DesignPackage.MinorVersion : 0,
                        PackageGuid = Guid.Parse(source.Publishing.DesignPackage.PackageGuid),
                        PackageName = source.Publishing.DesignPackage.PackageName,
                    } : null,
                    source.Publishing.AvailableWebTemplates != null && source.Publishing.AvailableWebTemplates.Length > 0 ?
                         (from awt in source.Publishing.AvailableWebTemplates
                          select new Model.AvailableWebTemplate
                          {
                              LanguageCode = awt.LanguageCodeSpecified ? awt.LanguageCode : 1033,
                              TemplateName = awt.TemplateName,
                          }) : null,
                    source.Publishing.PageLayouts != null && source.Publishing.PageLayouts.PageLayout != null && source.Publishing.PageLayouts.PageLayout.Length > 0 ?
                        (from pl in source.Publishing.PageLayouts.PageLayout
                         select new Model.PageLayout
                         {
                             IsDefault = pl.Path == source.Publishing.PageLayouts.Default,
                             Path = pl.Path,
                         }) : null
                    );
            }

            #endregion

            #region AddIns

            if (source.AddIns != null && source.AddIns.Length > 0)
            {
                result.AddIns.AddRange(
                     from addin in source.AddIns
                     select new Model.AddIn
                     {
                         PackagePath = addin.PackagePath,
                         Source = addin.Source.ToString(),
                     });
            }

            #endregion

            #region Providers

            // Translate Providers, if any
            if (source.Providers != null)
            {
                foreach (var provider in source.Providers)
                {
                    if (!String.IsNullOrEmpty(provider.HandlerType))
                    {
                        var handlerType = Type.GetType(provider.HandlerType, false);
                        if (handlerType != null)
                        {
                            result.ExtensibilityHandlers.Add(
                                new Model.ExtensibilityHandler
                                {
                                    Assembly = handlerType.Assembly.FullName,
                                    Type = handlerType.FullName,
                                    Configuration = provider.Configuration != null ? provider.Configuration.ToProviderConfiguration() : null,
                                    Enabled = provider.Enabled,
                                });
                        }
                    }
                }
            }

            #endregion

            return (result);
        }
    }

    internal static class V201508Extensions
    {
        public static V201508.Term[] FromModelTermsToSchemaTermsV201508(this TermCollection terms)
        {
            V201508.Term[] result = terms.Count > 0 ? (
                from term in terms
                select new V201508.Term
                {
                    ID = term.Id.ToString(),
                    Name = term.Name,
                    Description = term.Description,
                    Owner = term.Owner,
                    LanguageSpecified = term.Language.HasValue,
                    Language = term.Language.HasValue ? term.Language.Value : 1033,
                    IsAvailableForTagging = term.IsAvailableForTagging,
                    CustomSortOrder = term.CustomSortOrder,
                    Terms = term.Terms.Count > 0 ? new V201508.TermTerms { Items = term.Terms.FromModelTermsToSchemaTermsV201508() } : null,
                    CustomProperties = term.Properties.Count > 0 ?
                        (from p in term.Properties
                         select new V201508.StringDictionaryItem
                         {
                             Key = p.Key,
                             Value = p.Value
                         }).ToArray() : null,
                    LocalCustomProperties = term.LocalProperties.Count > 0 ?
                        (from p in term.LocalProperties
                         select new V201508.StringDictionaryItem
                         {
                             Key = p.Key,
                             Value = p.Value
                         }).ToArray() : null,
                    Labels = term.Labels.Count > 0 ?
                        (from l in term.Labels
                         select new V201508.TermLabelsLabel
                         {
                             Language = l.Language,
                             IsDefaultForLanguage = l.IsDefaultForLanguage,
                             Value = l.Value,
                         }).ToArray() : null,
                }).ToArray() : null;

            return (result);
        }

        public static List<Model.Term> FromSchemaTermsToModelTermsV201508(this V201508.Term[] terms)
        {
            List<Model.Term> result = new List<Model.Term>(
                from term in terms
                select new Model.Term(
                    !string.IsNullOrEmpty(term.ID) ? Guid.Parse(term.ID) : Guid.Empty,
                    term.Name,
                    term.LanguageSpecified ? term.Language : (int?)null,
                    (term.Terms != null && term.Terms.Items != null) ? term.Terms.Items.FromSchemaTermsToModelTermsV201508() : null,
                    term.Labels != null ?
                    (new List<Model.TermLabel>(
                        from label in term.Labels
                        select new Model.TermLabel
                        {
                            Language = label.Language,
                            Value = label.Value,
                            IsDefaultForLanguage = label.IsDefaultForLanguage
                        }
                    )) : null,
                    term.CustomProperties != null ? term.CustomProperties.ToDictionary(k => k.Key, v => v.Value) : null,
                    term.LocalCustomProperties != null ? term.LocalCustomProperties.ToDictionary(k => k.Key, v => v.Value) : null
                    )
                {
                    CustomSortOrder = term.CustomSortOrder,
                    IsAvailableForTagging = term.IsAvailableForTagging,
                    Owner = term.Owner,
                }
                );

            return (result);
        }

        public static V201508.CalendarType FromTemplateToSchemaCalendarTypeV201508(this Microsoft.SharePoint.Client.CalendarType calendarType)
        {
            switch (calendarType)
            {
                case Microsoft.SharePoint.Client.CalendarType.ChineseLunar:
                    return V201508.CalendarType.ChineseLunar;
                case Microsoft.SharePoint.Client.CalendarType.Gregorian:
                    return V201508.CalendarType.Gregorian;
                case Microsoft.SharePoint.Client.CalendarType.GregorianArabic:
                    return V201508.CalendarType.GregorianArabicCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianMEFrench:
                    return V201508.CalendarType.GregorianMiddleEastFrenchCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianXLITEnglish:
                    return V201508.CalendarType.GregorianTransliteratedEnglishCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianXLITFrench:
                    return V201508.CalendarType.GregorianTransliteratedFrenchCalendar;
                case Microsoft.SharePoint.Client.CalendarType.Hebrew:
                    return V201508.CalendarType.Hebrew;
                case Microsoft.SharePoint.Client.CalendarType.Hijri:
                    return V201508.CalendarType.Hijri;
                case Microsoft.SharePoint.Client.CalendarType.Japan:
                    return V201508.CalendarType.Japan;
                case Microsoft.SharePoint.Client.CalendarType.Korea:
                    return V201508.CalendarType.Korea;
                case Microsoft.SharePoint.Client.CalendarType.KoreaJapanLunar:
                    return V201508.CalendarType.KoreaandJapaneseLunar;
                case Microsoft.SharePoint.Client.CalendarType.SakaEra:
                    return V201508.CalendarType.SakaEra;
                case Microsoft.SharePoint.Client.CalendarType.Taiwan:
                    return V201508.CalendarType.Taiwan;
                case Microsoft.SharePoint.Client.CalendarType.Thai:
                    return V201508.CalendarType.Thai;
                case Microsoft.SharePoint.Client.CalendarType.UmAlQura:
                    return V201508.CalendarType.UmmalQura;
                case Microsoft.SharePoint.Client.CalendarType.None:
                default:
                    return V201508.CalendarType.None;
            }
        }

        public static Microsoft.SharePoint.Client.CalendarType FromSchemaToTemplateCalendarTypeV201508(this V201508.CalendarType calendarType)
        {
            switch (calendarType)
            {
                case V201508.CalendarType.ChineseLunar:
                    return Microsoft.SharePoint.Client.CalendarType.ChineseLunar;
                case V201508.CalendarType.Gregorian:
                    return Microsoft.SharePoint.Client.CalendarType.Gregorian;
                case V201508.CalendarType.GregorianArabicCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianArabic;
                case V201508.CalendarType.GregorianMiddleEastFrenchCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianMEFrench;
                case V201508.CalendarType.GregorianTransliteratedEnglishCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianXLITEnglish;
                case V201508.CalendarType.GregorianTransliteratedFrenchCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianXLITFrench;
                case V201508.CalendarType.Hebrew:
                    return Microsoft.SharePoint.Client.CalendarType.Hebrew;
                case V201508.CalendarType.Hijri:
                    return Microsoft.SharePoint.Client.CalendarType.Hijri;
                case V201508.CalendarType.Japan:
                    return Microsoft.SharePoint.Client.CalendarType.Japan;
                case V201508.CalendarType.Korea:
                    return Microsoft.SharePoint.Client.CalendarType.Korea;
                case V201508.CalendarType.KoreaandJapaneseLunar:
                    return Microsoft.SharePoint.Client.CalendarType.KoreaJapanLunar;
                case V201508.CalendarType.SakaEra:
                    return Microsoft.SharePoint.Client.CalendarType.SakaEra;
                case V201508.CalendarType.Taiwan:
                    return Microsoft.SharePoint.Client.CalendarType.Taiwan;
                case V201508.CalendarType.Thai:
                    return Microsoft.SharePoint.Client.CalendarType.Thai;
                case V201508.CalendarType.UmmalQura:
                    return Microsoft.SharePoint.Client.CalendarType.UmAlQura;
                case V201508.CalendarType.None:
                default:
                    return Microsoft.SharePoint.Client.CalendarType.None;
            }
        }

        public static V201508.WorkHour FromTemplateToSchemaWorkHourV201508(this Model.WorkHour workHour)
        {
            switch (workHour)
            {
                case Model.WorkHour.AM0100:
                    return V201508.WorkHour.Item100AM;
                case Model.WorkHour.AM0200:
                    return V201508.WorkHour.Item200AM;
                case Model.WorkHour.AM0300:
                    return V201508.WorkHour.Item300AM;
                case Model.WorkHour.AM0400:
                    return V201508.WorkHour.Item400AM;
                case Model.WorkHour.AM0500:
                    return V201508.WorkHour.Item500AM;
                case Model.WorkHour.AM0600:
                    return V201508.WorkHour.Item600AM;
                case Model.WorkHour.AM0700:
                    return V201508.WorkHour.Item700AM;
                case Model.WorkHour.AM0800:
                    return V201508.WorkHour.Item800AM;
                case Model.WorkHour.AM0900:
                    return V201508.WorkHour.Item900AM;
                case Model.WorkHour.AM1000:
                    return V201508.WorkHour.Item1000AM;
                case Model.WorkHour.AM1100:
                    return V201508.WorkHour.Item1100AM;
                case Model.WorkHour.AM1200:
                    return V201508.WorkHour.Item1200AM;
                case Model.WorkHour.PM0100:
                    return V201508.WorkHour.Item100PM;
                case Model.WorkHour.PM0200:
                    return V201508.WorkHour.Item200PM;
                case Model.WorkHour.PM0300:
                    return V201508.WorkHour.Item300PM;
                case Model.WorkHour.PM0400:
                    return V201508.WorkHour.Item400PM;
                case Model.WorkHour.PM0500:
                    return V201508.WorkHour.Item500PM;
                case Model.WorkHour.PM0600:
                    return V201508.WorkHour.Item600PM;
                case Model.WorkHour.PM0700:
                    return V201508.WorkHour.Item700PM;
                case Model.WorkHour.PM0800:
                    return V201508.WorkHour.Item800PM;
                case Model.WorkHour.PM0900:
                    return V201508.WorkHour.Item900PM;
                case Model.WorkHour.PM1000:
                    return V201508.WorkHour.Item1000PM;
                case Model.WorkHour.PM1100:
                    return V201508.WorkHour.Item1100PM;
                case Model.WorkHour.PM1200:
                    return V201508.WorkHour.Item1200PM;
                default:
                    return V201508.WorkHour.Item100AM;
            }
        }

        public static Model.WorkHour FromSchemaToTemplateWorkHourV201508(this V201508.WorkHour workHour)
        {
            switch (workHour)
            {
                case V201508.WorkHour.Item100AM:
                    return Model.WorkHour.AM0100;
                case V201508.WorkHour.Item200AM:
                    return Model.WorkHour.AM0200;
                case V201508.WorkHour.Item300AM:
                    return Model.WorkHour.AM0300;
                case V201508.WorkHour.Item400AM:
                    return Model.WorkHour.AM0400;
                case V201508.WorkHour.Item500AM:
                    return Model.WorkHour.AM0500;
                case V201508.WorkHour.Item600AM:
                    return Model.WorkHour.AM0600;
                case V201508.WorkHour.Item700AM:
                    return Model.WorkHour.AM0700;
                case V201508.WorkHour.Item800AM:
                    return Model.WorkHour.AM0800;
                case V201508.WorkHour.Item900AM:
                    return Model.WorkHour.AM0900;
                case V201508.WorkHour.Item1000AM:
                    return Model.WorkHour.AM1000;
                case V201508.WorkHour.Item1100AM:
                    return Model.WorkHour.AM1100;
                case V201508.WorkHour.Item1200AM:
                    return Model.WorkHour.AM1200;
                case V201508.WorkHour.Item100PM:
                    return Model.WorkHour.PM0100;
                case V201508.WorkHour.Item200PM:
                    return Model.WorkHour.PM0200;
                case V201508.WorkHour.Item300PM:
                    return Model.WorkHour.PM0300;
                case V201508.WorkHour.Item400PM:
                    return Model.WorkHour.PM0400;
                case V201508.WorkHour.Item500PM:
                    return Model.WorkHour.PM0500;
                case V201508.WorkHour.Item600PM:
                    return Model.WorkHour.PM0600;
                case V201508.WorkHour.Item700PM:
                    return Model.WorkHour.PM0700;
                case V201508.WorkHour.Item800PM:
                    return Model.WorkHour.PM0800;
                case V201508.WorkHour.Item900PM:
                    return Model.WorkHour.PM0900;
                case V201508.WorkHour.Item1000PM:
                    return Model.WorkHour.PM1000;
                case V201508.WorkHour.Item1100PM:
                    return Model.WorkHour.PM1100;
                case V201508.WorkHour.Item1200PM:
                    return Model.WorkHour.PM1200;
                default:
                    return Model.WorkHour.AM0100;
            }
        }

        public static V201508.AuditSettingsAudit[] FromTemplateToSchemaAuditsV201508(this Microsoft.SharePoint.Client.AuditMaskType audits)
        {
            List<V201508.AuditSettingsAudit> result = new List<AuditSettingsAudit>();
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.All))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.All });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.CheckIn))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.CheckIn });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.CheckOut))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.CheckOut });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.ChildDelete))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.ChildDelete });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Copy))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Copy });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Move))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Move });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.None))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.None });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.ObjectDelete))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.ObjectDelete });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.ProfileChange))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.ProfileChange });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.SchemaChange))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.SchemaChange });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Search))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Search });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.SecurityChange))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.SecurityChange });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Undelete))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Undelete });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Update))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Update });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.View))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.View });
            }
            if (audits.HasFlag(Microsoft.SharePoint.Client.AuditMaskType.Workflow))
            {
                result.Add(new AuditSettingsAudit { AuditFlag = AuditSettingsAuditAuditFlag.Workflow });
            }

            return result.ToArray();
        }

        public static Model.ObjectSecurity FromSchemaToTemplateObjectSecurityV201508(this V201508.ObjectSecurity objectSecurity)
        {
            return ((objectSecurity != null && objectSecurity.BreakRoleInheritance != null) ?
                new Model.ObjectSecurity(
                    objectSecurity.BreakRoleInheritance.RoleAssignment != null ?
                        (from ra in objectSecurity.BreakRoleInheritance.RoleAssignment
                         select new Model.RoleAssignment
                         {
                             Principal = ra.Principal,
                             RoleDefinition = ra.RoleDefinition,
                         }) : null
                    )
                {
                    ClearSubscopes = objectSecurity.BreakRoleInheritance.ClearSubscopes,
                    CopyRoleAssignments = objectSecurity.BreakRoleInheritance.CopyRoleAssignments,
                } : null);
        }

        public static V201508.ObjectSecurity FromTemplateToSchemaObjectSecurityV201508(this Model.ObjectSecurity objectSecurity)
        {
            return ((objectSecurity != null && (objectSecurity.ClearSubscopes == true || objectSecurity.CopyRoleAssignments == true || objectSecurity.RoleAssignments.Count > 0)) ?
                new V201508.ObjectSecurity
                {
                    BreakRoleInheritance = new V201508.ObjectSecurityBreakRoleInheritance
                    {
                        ClearSubscopes = objectSecurity.ClearSubscopes,
                        CopyRoleAssignments = objectSecurity.CopyRoleAssignments,
                        RoleAssignment = (objectSecurity.RoleAssignments != null && objectSecurity.RoleAssignments.Count > 0) ?
                            (from ra in objectSecurity.RoleAssignments
                             select new V201508.RoleAssignment
                             {
                                 Principal = ra.Principal,
                                 RoleDefinition = ra.RoleDefinition,
                             }).ToArray() : null,
                    }
                } : null);
        }
    }
}

