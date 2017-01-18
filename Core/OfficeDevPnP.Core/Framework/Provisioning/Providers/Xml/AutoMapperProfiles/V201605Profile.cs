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
    public class V201605Profile : Profile
    {
        public V201605Profile()
        {
            // To keep in account
            // - *Specified -> we can try to define a generic AfterMap handler
            // - Custom conversion for specific properties
            // - Default values for missing properties
            // - Collections with custom properties, like ListInstanceViews

            CreateMap<Model.ProvisioningTemplate, Schema.ProvisioningTemplate>().ReverseMap();
            // TODO: Double check Version conversion from double to decimal

            // Template Properties
            CreateMap<Dictionary<String, String>, Schema.StringDictionaryItem[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.StringDictionaryItem>(i => i.Key, i => i.Value));
            // TODO: Manually do the Reverse path

            // Localization
            CreateMap<Model.Localization, Schema.LocalizationsLocalization>().ReverseMap();

            // Property Bags
            CreateMap<Model.PropertyBagEntry, Schema.PropertyBagEntry>().ReverseMap();
            // TODO: OverwriteSpecified

            // Web Settings
            CreateMap<Model.WebSettings, Schema.WebSettings>().ReverseMap();
            // TODO: NoCrawlSpecified

            // Regional Settings
            CreateMap<Model.RegionalSettings, Schema.RegionalSettings>().ReverseMap();
            // TODO: AdjustHijriDaysSpecified, AlternateCalendarTypeSpecified, CalendarTypeSpecified, 
            // CollationSpecified, FirstDayOfWeekSpecified, FirstWeekOfYearSpecified, LocaleIdSpecified, 
            // ShowWeeksSpecified, Time24Specified, WorkDayEndHourSpecified, WorkDaysSpecified, 
            // WorkDayStartHourSpecified
            // TODO: FromTemplateToSchemaCalendarTypeV201605, FromTemplateToSchemaWorkHourV201605

            // Supported UI Languages
            CreateMap<Model.SupportedUILanguage, Schema.SupportedUILanguagesSupportedUILanguage>().ReverseMap();

            // Audit Settings
            CreateMap<Model.AuditSettings, Schema.AuditSettings>().ReverseMap();
            // TODO: AuditLogTrimmingRetentionSpecified, TrimAuditLogSpecified
            // TODO: FromTemplateToSchemaAuditsV201605

            // Site Security
            CreateMap<Model.SiteSecurity, Schema.Security>().ReverseMap();
            CreateMap<Model.User, Schema.User>().ReverseMap();
            CreateMap<Model.SiteGroup, Schema.SiteGroup>().ReverseMap();
            // TODO: AllowMembersEditMembershipSpecified, AllowRequestToJoinLeaveSpecified, 
            // AutoAcceptRequestToJoinLeaveSpecified, OnlyAllowMembersViewMembershipSpecified
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
                    .ResolveUsing(new FromStringToDocumentTemplateResolver<Schema.ContentType, Schema.ContentTypeDocumentTemplate>()));
            // TODO: Manually do the Reverse path
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
                        Schema.DocumentSetTemplate, Schema.DocumentSetFieldRef[]>()));
            // TODO: Manually do the Reverse path
            CreateMap<Model.DefaultDocument, Schema.DocumentSetTemplateDefaultDocument>().ReverseMap();

            // List Instances
            CreateMap<Model.ListInstance, Schema.ListInstance>().ReverseMap();
            // TODO: DraftVersionVisibilitySpecified, MinorVersionLimitSpecified, MaxVersionLimitSpecified
            CreateMap<Model.ContentTypeBinding, Schema.ContentTypeBinding>().ReverseMap();
            CreateMap<Model.ViewCollection, Schema.ListInstanceViews>()
                .ConvertUsing(new FromCollectionToXmlElementsConverter<Model.ViewCollection, Schema.ListInstanceViews>());
            // TODO: RemoveExistingViews
            // TODO: Manually do the Reverse path
            // The Fields have already been defined before for Site Fields
            CreateMap<Dictionary<String, String>, Schema.FieldDefault[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.FieldDefault>(i => i.FieldName, i => i.Value));
            // TODO: Manually do the Reverse path
            CreateMap<Model.FieldRef, Schema.ListInstanceFieldRef>().ReverseMap();
            CreateMap<Model.DataRow, Schema.ListInstanceDataRow>().ReverseMap();
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            CreateMap<Dictionary<String, String>, Schema.DataValue[]>()
                .ConvertUsing(new FromDictionaryToArrayConverter<String, String, Schema.DataValue>(i => i.FieldName, i => i.Value));
            // TODO: FromTemplateToSchemaObjectSecurityV201605
            // TODO: FromTemplateToSchemaFolderV201605
            // TODO: FromBasePermissionsToStringV201605
            // TODO: Manually do the Reverse path
            CreateMap<Model.CustomAction, Schema.CustomAction>().ReverseMap();
            // TODO: RegistrationTypeSpecified, SequenceSpecified
            // TODO: CustomActionCommandUIExtension

            // Features
            CreateMap<Model.Feature, Schema.Feature>().ReverseMap();

            // Custom Actions
            // Nothing to do here, because we already defined CustomActions at List Instances level

            // Files
            CreateMap<Model.File, Schema.File>().ReverseMap();
            // TODO: LevelSpecified
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
            CreateMap<Model.TermGroup, Schema.TermGroup>().ReverseMap();
            // TODO: SiteCollectionTermGroupSpecified, 
            // Contributors and Managers have already been defined previously
            CreateMap<Model.TermSet, Schema.TermSet>().ReverseMap();
            // TODO: LanguageSpecified
            // Properties have already been defined previously

            // Composed Looks
            CreateMap<Model.ComposedLook, Schema.ComposedLook>().ReverseMap();
            // TODO: Post-validation -> put to NULL an empty composed look

            // Workflows
            CreateMap<Model.Workflows, Schema.Workflows>().ReverseMap();
            CreateMap<Model.WorkflowDefinition, Schema.WorkflowsWorkflowDefinition>().ReverseMap();
            // TODO: PublishedSpecified, RequiresAssociationFormSpecified, 
            // RequiresInitiationFormSpecified, RestrictToTypeSpecified
            CreateMap<Model.WorkflowSubscription, Schema.WorkflowsWorkflowSubscription>().ReverseMap();
            // TODO: ManualStartBypassesActivationLimitSpecified

            // Search Settings
            // TODO: Handle WebSearchSettings, SiteSearchSettings to XmlElement

            // Publishing
            CreateMap<Model.Publishing, Schema.Publishing>().ReverseMap();
            CreateMap<Model.AvailableWebTemplate, Schema.PublishingWebTemplate>().ReverseMap();
            // TODO: LanguageCodeSpecified
            CreateMap<Model.DesignPackage, Schema.PublishingDesignPackage>().ReverseMap();
            // TODO: MajorVersionSpecified, MinorVersionSpecified
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
