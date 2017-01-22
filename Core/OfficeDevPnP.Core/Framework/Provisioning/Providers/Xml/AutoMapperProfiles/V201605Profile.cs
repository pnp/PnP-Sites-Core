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
    /// <summary>
    /// AutoMapper Profile class for schema V201605, inheriting from V201512 (the previous schema version)
    /// </summary>
    public class V201605Profile : Profile
    {
        public V201605Profile()
        {
            // To keep in account
            // - *Specified -> we can try to define a generic AfterMap handler
            // - Custom conversion for specific properties
            // - Default values for missing properties
            // - Collections with custom properties, like ListInstanceViews

            // Define double-way mapping for Provisioning Template
            // Special properties: VersionSpecified
            CreateMap<Model.ProvisioningTemplate, Schema.ProvisioningTemplate>()
                .HandleSpecifiedProperties()
                .ForMember(pt => pt.Version, opt => opt.ResolveUsing(new FromDoubleToDecimalResolver(), src => src.Version))
                .ReverseMap()
                .HandleSpecifiedProperties()
                .ForMember(pt => pt.Version, opt => opt.ResolveUsing(new FromDecimalToDoubleResolver(), src => src.Version));

            // Template Properties
            CreateMap<Dictionary<String, String>, Schema.StringDictionaryItem[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.StringDictionaryItem>(i => i.Key, i => i.Value));
            // TODO: Manually do the Reverse path

            // Localization
            CreateMap<Model.Localization, Schema.LocalizationsLocalization>().ReverseMap();

            // Property Bags
            // Special properties: OverwriteSpecified
            CreateMap<Model.PropertyBagEntry, Schema.PropertyBagEntry>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();

            // Web Settings
            // Special properties: NoCrawlSpecified
            CreateMap<Model.WebSettings, Schema.WebSettings>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();

            // Regional Settings
            // Special properties: AdjustHijriDaysSpecified, AlternateCalendarTypeSpecified, CalendarTypeSpecified, 
            // CollationSpecified, FirstDayOfWeekSpecified, FirstWeekOfYearSpecified, LocaleIdSpecified, 
            // ShowWeeksSpecified, Time24Specified, WorkDayEndHourSpecified, WorkDaysSpecified, 
            // WorkDayStartHourSpecified
            CreateMap<Model.RegionalSettings, Schema.RegionalSettings>()
                .HandleSpecifiedProperties()
                .ForMember(dest => dest.AlternateCalendarType, opt => opt.ResolveUsing(null, src => src.AlternateCalendarType)) // TODO: Create a custom resolver (see FromTemplateToSchemaCalendarTypeV201605)
                .ForMember(dest => dest.WorkDayStartHour, opt => opt.ResolveUsing(null, src => src.WorkDayStartHour)) // TODO: Create a custom resolver (see FromTemplateToSchemaWorkHourV201605)
                .ForMember(dest => dest.WorkDayEndHour, opt => opt.ResolveUsing(null, src => src.WorkDayEndHour)) // TODO: Create a custom resolver (the same as the one before)
                .ReverseMap()
                .HandleSpecifiedProperties()
                .ForMember(dest => dest.AlternateCalendarType, opt => opt.ResolveUsing(null, src => src.AlternateCalendarType)) // TODO: Create a custom resolver (see FromTemplateToSchemaCalendarTypeV201605)
                .ForMember(dest => dest.WorkDayStartHour, opt => opt.ResolveUsing(null, src => src.WorkDayStartHour)) // TODO: Create a custom resolver (see FromTemplateToSchemaWorkHourV201605)
                .ForMember(dest => dest.WorkDayEndHour, opt => opt.ResolveUsing(null, src => src.WorkDayEndHour)); // TODO: Create a custom resolver (the same as the one before)

            // Supported UI Languages
            CreateMap<Model.SupportedUILanguage, Schema.SupportedUILanguagesSupportedUILanguage>().ReverseMap();

            // Audit Settings
            // Special properties: AuditLogTrimmingRetentionSpecified, TrimAuditLogSpecified
            CreateMap<Model.AuditSettings, Schema.AuditSettings>()
                .HandleSpecifiedProperties()
                .ForMember(dest => dest.Audit, opt => opt.ResolveUsing(null, src => src.AuditFlags)) // TODO: Create a custom resolver (see FromTemplateToSchemaAuditsV201605)
                .ReverseMap()
                .HandleSpecifiedProperties()
                .ForMember(dest => dest.AuditFlags, opt => opt.ResolveUsing(null, src => src.Audit)); // TODO: Create a custom resolver (see FromTemplateToSchemaAuditsV201605)

            // Site Security
            CreateMap<Model.SiteSecurity, Schema.Security>().ReverseMap();
            CreateMap<Model.User, Schema.User>().ReverseMap();
            // Special properties: AllowMembersEditMembershipSpecified, AllowRequestToJoinLeaveSpecified, 
            // AutoAcceptRequestToJoinLeaveSpecified, OnlyAllowMembersViewMembershipSpecified
            CreateMap<Model.SiteGroup, Schema.SiteGroup>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            CreateMap<Model.RoleAssignment, Schema.RoleAssignment>().ReverseMap();
            CreateMap<Model.RoleDefinition, Schema.RoleDefinition>().ReverseMap();

            // Navigation
            CreateMap<Model.Navigation, Schema.Navigation>().ReverseMap();
            CreateMap<Model.GlobalNavigation, Schema.NavigationGlobalNavigation>().ReverseMap();
            CreateMap<Model.CurrentNavigation, Schema.NavigationCurrentNavigation>().ReverseMap();
            CreateMap<Model.ManagedNavigation, Schema.ManagedNavigation>().ReverseMap();
            CreateMap<Model.StructuralNavigation, Schema.StructuralNavigation>().ReverseMap();
            CreateMap<Model.NavigationNode, Schema.NavigationNode>().ReverseMap();

            // Site Fields
            CreateMap<Model.FieldCollection, Schema.ProvisioningTemplateSiteFields>()
                .ConvertUsing(new FromCollectionToXmlElementsConverter<Model.FieldCollection, Schema.ProvisioningTemplateSiteFields>());
            // TODO: Manually do the Reverse path

            // Content Types
            CreateMap<Model.ContentType, Schema.ContentType>()
                .ForMember(dest => dest.DocumentTemplate, opt => opt
                    .ResolveUsing(new FromStringToDocumentTemplateResolver<Schema.ContentType, Schema.ContentTypeDocumentTemplate>()))
                .ReverseMap()
                .ForMember(dest => dest.DocumentTemplate, opt => opt
                    .ResolveUsing(null)); // TODO: Create the custom converter
            CreateMap<Model.FieldRef, Schema.ContentTypeFieldRef>().ReverseMap();
            CreateMap<Model.DocumentSetTemplate, Schema.DocumentSetTemplate>()
                .ForMember(dest => dest.AllowedContentTypes, opt => opt
                    .ResolveUsing(new FromListToTypedArrayResolver<Model.DocumentSetTemplate,
                        Schema.DocumentSetTemplate, Schema.DocumentSetTemplateAllowedContentType[]>()))
                .ForMember(dest => dest.SharedFields, opt => opt
                    .ResolveUsing(new FromListToTypedArrayResolver<Model.DocumentSetTemplate,
                        Schema.DocumentSetTemplate, Schema.DocumentSetFieldRef[]>()))
                .ForMember(dest => dest.WelcomePageFields, opt => opt
                    .ResolveUsing(new FromListToTypedArrayResolver<Model.DocumentSetTemplate,
                        Schema.DocumentSetTemplate, Schema.DocumentSetFieldRef[]>()))
                .ReverseMap()
                .ForMember(dest => dest.AllowedContentTypes, opt => opt
                    .ResolveUsing(null)) // TODO: Create the custom converter
                .ForMember(dest => dest.SharedFields, opt => opt
                    .ResolveUsing(null)) // TODO: Create the custom converter
                .ForMember(dest => dest.WelcomePageFields, opt => opt
                    .ResolveUsing(null)); // TODO: Create the custom converter
            CreateMap<Model.DefaultDocument, Schema.DocumentSetTemplateDefaultDocument>().ReverseMap();

            // List Instances
            // Special properties: DraftVersionVisibilitySpecified, MinorVersionLimitSpecified, MaxVersionLimitSpecified
            CreateMap<Model.ListInstance, Schema.ListInstance>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            CreateMap<Model.ContentTypeBinding, Schema.ContentTypeBinding>().ReverseMap();
            CreateMap<Model.ViewCollection, Schema.ListInstanceViews>()
                .ConvertUsing(new FromCollectionToXmlElementsConverter<Model.ViewCollection, Schema.ListInstanceViews>());
            // TODO: Handle the RemoveExistingViews attribute
            // TODO: Manually do the Reverse path
            // The Fields have already been defined before for Site Fields
            CreateMap<Dictionary<String, String>, Schema.FieldDefault[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.FieldDefault>(i => i.FieldName, i => i.Value));
            // TODO: Manually do the Reverse path
            CreateMap<Model.FieldRef, Schema.ListInstanceFieldRef>().ReverseMap();
            CreateMap<Model.DataRow, Schema.ListInstanceDataRow>()
                .ForMember(dest => dest.Security, opt => opt
                    .ResolveUsing(null)) // TODO: Create the custom resolver (see FromTemplateToSchemaObjectSecurityV201605)
                .ReverseMap()
                .ForMember(dest => dest.Security, opt => opt
                    .ResolveUsing(null)); // TODO: Create the custom resolver (see FromTemplateToSchemaObjectSecurityV201605)
            CreateMap<Dictionary<String, String>, Schema.DataValue[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.DataValue>(i => i.FieldName, i => i.Value));
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            // TODO: FromTemplateToSchemaFolderV201605
            // TODO: FromBasePermissionsToStringV201605
            // TODO: Manually do the Reverse path
            // Special properties: RegistrationTypeSpecified, SequenceSpecified
            CreateMap<Model.CustomAction, Schema.CustomAction>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            // TODO: CustomActionCommandUIExtension

            // Features
            CreateMap<Model.Feature, Schema.Feature>().ReverseMap();

            // Custom Actions
            // Nothing to do here, because we already defined CustomActions at List Instances level

            // Files
            // Special properties: LevelSpecified
            CreateMap<Model.File, Schema.File>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            CreateMap<Model.WebPart, Schema.WebPartPageWebPart>().ReverseMap();
            // TODO: Custom handling of Contents for WebPart
            // The Properties are already defined at the beginning of the Profile constructor

            // Pages
            CreateMap<Model.Page, Schema.Page>().ReverseMap();
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            CreateMap<Model.WebPart, Schema.WikiPageWebPart>().ReverseMap();
            // TODO: Custom handling of Contents for WebPart
            CreateMap<Dictionary<String, String>, Schema.BaseFieldValue[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.BaseFieldValue>(i => i.FieldName, i => i.Value));
            // TODO: Manually do the Reverse path

            // Taxonomy
            // Special properties: SiteCollectionTermGroupSpecified
            CreateMap<Model.TermGroup, Schema.TermGroup>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            // Contributors and Managers have already been defined previously
            // Special properties: LanguageSpecified
            CreateMap<Model.TermSet, Schema.TermSet>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            // Properties have already been defined previously

            // Composed Looks
            CreateMap<Model.ComposedLook, Schema.ComposedLook>().ReverseMap();
            // TODO: Post-validation -> put to NULL an empty composed look

            // Workflows
            CreateMap<Model.Workflows, Schema.Workflows>().ReverseMap();
            // Special properties: PublishedSpecified, RequiresAssociationFormSpecified, 
            // RequiresInitiationFormSpecified, RestrictToTypeSpecified
            CreateMap<Model.WorkflowDefinition, Schema.WorkflowsWorkflowDefinition>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            // Special properties: ManualStartBypassesActivationLimitSpecified
            CreateMap<Model.WorkflowSubscription, Schema.WorkflowsWorkflowSubscription>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();

            // Search Settings
            // TODO: Handle WebSearchSettings, SiteSearchSettings to XmlElement

            // Publishing
            CreateMap<Model.Publishing, Schema.Publishing>().ReverseMap();
            // Special properties: LanguageCodeSpecified
            CreateMap<Model.AvailableWebTemplate, Schema.PublishingWebTemplate>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            // Special properties: MajorVersionSpecified, MinorVersionSpecified
            CreateMap<Model.DesignPackage, Schema.PublishingDesignPackage>()
                .HandleSpecifiedProperties()
                .ReverseMap()
                .HandleSpecifiedProperties();
            CreateMap<Model.PageLayout, Schema.PublishingPageLayoutsPageLayout>().ReverseMap();
            // TODO: Be careful about the PublishingPageLayouts container type
            // TODO: Handle the default value for the default page layout

            // AddIns
            CreateMap<Model.AddIn, Schema.AddInsAddin>().ReverseMap();

            // Providers
            // TODO: Remember to add both ExtensibilityHandlers and templates Providers
        }
    }
}
