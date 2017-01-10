using AutoMapper;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Schema = OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperProfiles
{
    public class FromDomainModelToV201605Profile : Profile
    {
        public FromDomainModelToV201605Profile()
        {
            // To keep in account
            // - *Specified -> we can try to define a generic AfterMap handler
            // - Custom conversion for specific properties
            // - Default values for missing properties
            // - Collections with custom properties, like ListInstanceViews

            CreateMap<Model.ProvisioningTemplate, Schema.ProvisioningTemplate>();
            // TODO: Double check Version conversion from double to decimal

            // Template Properties
            CreateMap<Dictionary<String, String>, Schema.StringDictionaryItem[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.StringDictionaryItem>(i => i.Key, i => i.Value));

            // Localization
            CreateMap<Model.Localization, Schema.LocalizationsLocalization>();

            // Property Bags
            CreateMap<Model.PropertyBagEntry, Schema.PropertyBagEntry>();
            // TODO: OverwriteSpecified

            // Web Settings
            CreateMap<Model.WebSettings, Schema.WebSettings>();
            // TODO: NoCrawlSpecified

            // Regional Settings
            CreateMap<Model.RegionalSettings, Schema.RegionalSettings>();
            // TODO: AdjustHijriDaysSpecified, AlternateCalendarTypeSpecified, CalendarTypeSpecified, 
            // CollationSpecified, FirstDayOfWeekSpecified, FirstWeekOfYearSpecified, LocaleIdSpecified, 
            // ShowWeeksSpecified, Time24Specified, WorkDayEndHourSpecified, WorkDaysSpecified, 
            // WorkDayStartHourSpecified
            // TODO: FromTemplateToSchemaCalendarTypeV201605, FromTemplateToSchemaWorkHourV201605

            // Supported UI Languages
            CreateMap<Model.SupportedUILanguage, Schema.SupportedUILanguagesSupportedUILanguage>();

            // Audit Settings
            CreateMap<Model.AuditSettings, Schema.AuditSettings>();
            // TODO: AuditLogTrimmingRetentionSpecified, TrimAuditLogSpecified
            // TODO: FromTemplateToSchemaAuditsV201605

            // Site Security
            CreateMap<Model.SiteSecurity, Schema.Security>();
            CreateMap<Model.User, Schema.User>();
            CreateMap<Model.SiteGroup, Schema.SiteGroup>();
            // TODO: AllowMembersEditMembershipSpecified, AllowRequestToJoinLeaveSpecified, 
            // AutoAcceptRequestToJoinLeaveSpecified, OnlyAllowMembersViewMembershipSpecified
            CreateMap<Model.RoleAssignment, Schema.RoleAssignment>();
            CreateMap<Model.RoleDefinition, Schema.RoleDefinition>();

            // Navigation
            CreateMap<Model.Navigation, Schema.Navigation>();
            CreateMap<Model.GlobalNavigation, Schema.NavigationGlobalNavigation>();
            CreateMap<Model.CurrentNavigation, Schema.NavigationCurrentNavigation>();
            CreateMap<Model.ManagedNavigation, Schema.ManagedNavigation>();
            CreateMap<Model.StructuralNavigation, Schema.StructuralNavigation>();
            CreateMap<Model.NavigationNode, Schema.NavigationNode>();

            // Site Fields
            CreateMap<Model.FieldCollection, Schema.ProvisioningTemplateSiteFields>()
                .ConvertUsing(new FromCollectionToXmlElementsConverter<Model.FieldCollection, Schema.ProvisioningTemplateSiteFields>());

            // Content Types
            CreateMap<Model.ContentType, Schema.ContentType>()
                .ForMember(dest => dest.DocumentTemplate, opt => opt
                    .ResolveUsing(new FromStringToDocumentTemplateResolver<Schema.ContentType, Schema.ContentTypeDocumentTemplate>()));
            CreateMap<Model.FieldRef, Schema.ContentTypeFieldRef>();
            CreateMap<Model.DocumentSetTemplate, Schema.DocumentSetTemplate>()
                .ForMember(dest => dest.AllowedContentTypes, opt => opt
                    .ResolveUsing(new FromListToTypedArrayResolver<Model.DocumentSetTemplate,
                        Schema.DocumentSetTemplate, Schema.DocumentSetTemplateAllowedContentType[]>()))
                .ForMember(dest => dest.SharedFields, opt => opt
                    .ResolveUsing(new FromListToTypedArrayResolver<Model.DocumentSetTemplate,
                        Schema.DocumentSetTemplate, Schema.DocumentSetFieldRef[]>()))
                .ForMember(dest => dest.WelcomePageFields, opt => opt
                    .ResolveUsing(new FromListToTypedArrayResolver<Model.DocumentSetTemplate,
                        Schema.DocumentSetTemplate, Schema.DocumentSetFieldRef[]>()));
            CreateMap<Model.DefaultDocument, Schema.DocumentSetTemplateDefaultDocument>();

            // List Instances
            CreateMap<Model.ListInstance, Schema.ListInstance>();
            // TODO: DraftVersionVisibilitySpecified, MinorVersionLimitSpecified, MaxVersionLimitSpecified
            CreateMap<Model.ContentTypeBinding, Schema.ContentTypeBinding>();
            CreateMap<Model.ViewCollection, Schema.ListInstanceViews>()
                .ConvertUsing(new FromCollectionToXmlElementsConverter<Model.ViewCollection, Schema.ListInstanceViews>());
            // TODO: RemoveExistingViews
            // The Fields have already been defined before for Site Fields
            CreateMap<Dictionary<String, String>, Schema.FieldDefault[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.FieldDefault>(i => i.FieldName, i => i.Value));
            CreateMap<Model.FieldRef, Schema.ListInstanceFieldRef>();
            CreateMap<Model.DataRow, Schema.ListInstanceDataRow>();
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            CreateMap<Dictionary<String, String>, Schema.DataValue[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.DataValue>(i => i.FieldName, i => i.Value));
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            // TODO: FromTemplateToSchemaFolderV201605
            // TODO: FromBasePermissionsToStringV201605
            CreateMap<Model.CustomAction, Schema.CustomAction>();
            // TODO: RegistrationTypeSpecified, SequenceSpecified
            // TODO: CustomActionCommandUIExtension

            // Features
            CreateMap<Model.Feature, Schema.Feature>();

            // Custom Actions
            // Nothing to do here, because we already defined CustomActions at List Instances level

            // Files
            CreateMap<Model.File, Schema.File>();
            // TODO: LevelSpecified
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            CreateMap<Model.WebPart, Schema.WebPartPageWebPart>();
            // TODO: Custom handling of Contents for WebPart
            // The Properties are already defined at the beginning of the Profile constructor

            // Pages
            CreateMap<Model.Page, Schema.Page>();
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            CreateMap<Model.WebPart, Schema.WikiPageWebPart>();
            // TODO: Custom handling of Contents for WebPart
            CreateMap<Dictionary<String, String>, Schema.BaseFieldValue[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.BaseFieldValue>(i => i.FieldName, i => i.Value));

            // Taxonomy
            CreateMap<Model.TermGroup, Schema.TermGroup>();
            // TODO: SiteCollectionTermGroupSpecified, 
            // Contributors and Managers have already been defined previously
            CreateMap<Model.TermSet, Schema.TermSet>();
            // TODO: LanguageSpecified
            // Properties have already been defined previously

            // Composed Looks
            CreateMap<Model.ComposedLook, Schema.ComposedLook>();
            // TODO: Post-validation -> put to NULL an empty composed look

            // Workflows
            CreateMap<Model.Workflows, Schema.Workflows>();
            CreateMap<Model.WorkflowDefinition, Schema.WorkflowsWorkflowDefinition>();
            // TODO: PublishedSpecified, RequiresAssociationFormSpecified, 
            // RequiresInitiationFormSpecified, RestrictToTypeSpecified
            CreateMap<Model.WorkflowSubscription, Schema.WorkflowsWorkflowSubscription>();
            // TODO: ManualStartBypassesActivationLimitSpecified

            // Search Settings
            // TODO: Handle WebSearchSettings, SiteSearchSettings to XmlElement

            // Publishing
            CreateMap<Model.Publishing, Schema.Publishing>();
            CreateMap<Model.AvailableWebTemplate, Schema.PublishingWebTemplate>();
            // TODO: LanguageCodeSpecified
            CreateMap<Model.DesignPackage, Schema.PublishingDesignPackage>();
            // TODO: MajorVersionSpecified, MinorVersionSpecified
            CreateMap<Model.PageLayout, Schema.PublishingPageLayoutsPageLayout>();
            // TODO: Be careful about the PublishingPageLayouts container type
            // TODO: Handle the default value for the default page layout

            // AddIns
            CreateMap<Model.AddIn, Schema.AddInsAddin>();

            // Providers
            // TODO: Remember to add both ExtensibilityHandlers and templates Providers
        }
    }
}
