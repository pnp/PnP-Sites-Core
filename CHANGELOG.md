# OfficeDevPnP.Sites.Core Changelog

*Please do not commit changes to this file, it is maintained by the repo owner.*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/).
## [2.26.1805.0 - Unreleased]

### Added
- Added CommentsOnSitePagesDisabled property on web settings element in the provisioning engine.
- Added support for StorageEntities to the Tenant element in the Provisioning Engine. The user applying the template needs appropriate access rights to the tenant scoped App Catalog.
- Added SiteScripts and SiteDesigns elements to the Tenant element in the Provisioning Engine. The user applying the template needs to be tenant administrator.
- Added HubSiteUrl to the WebSettings element for the Provisioning Engine. The user applying the template needs to be tenant administrator.
- Added {SiteScriptId:[script title]} and {SiteDesignId:[design title]} tokens to the provisioning engine. This will only work if the user applying the template is tenant administrator.
- Added {StorageEntityValue:[key]} token to retrieve values from tenant level or (when applicable) site collection level. If a key is present at site collection level this value will take preference over the one from tenant level, following the behavior of the CSOM APIs.

### Changed

### Deprecated

## [2.25.1804.0 - April 2018 release]

### Added

- Added async external sharing extension methods [baywet]
- Added ProvisionFieldsToSubWebs option to ProvisioningTemplateApplyingInformation class [jensotto]
- Addition of PnPCore.Tests project for testing of the PnPCore .Net Standard 2.0 library
- Added Scope parameter to ALM Manager methods allowing you to perform application lifecycle management tasks to the site collection scoped app catalog.

### Changed

- Added support for CDN Elements in Provisioning Engine
- Support for FullBleed configuration for adding web parts in "Full Width column" section [OliverZeiser]
- Improvements to ExecuteQueryRetryAsync [OliverZeiser, biste5]
- Improvements to support provisioning engine to be called from non console applications
- Better support for async methods, avoiding deadlocks
- Updated spelling across various files [fowl2]
- Refactored ObjectListHandler [stevebeauge]

### Deprecated

## [2.24.1803.0 - March 2018 release]

### Added

- Added ExecuteQueryRetryAsync method [baywet and SharePointRadi]
- Added EnsureLabel extension method to the taxonomy extensions [paulpascha]
- Added SetDefaultContentType extension methods on List objects. Notice that this method behaves different from the deprecated SetDefaultContentTypeToList method. See the Deprecated section.
- Added AliasExistsAsync extension method to verify if an Office 365 Group alias is available for use
- Added support for taxonomy fields in DataRows at the provisioning engine level. [jensotto]
- Added support for updating owners and members of an Office 365 Group.
- Added support for TermStore DefaultLanguage when retrieving or adding a term. [stevebeauge]
- Added support for getting apps by title [gautamdsheth]
- Added .NET 2.0 Standard project to allow cross-platform use of the PnP Sites Core library

### Changed

- Improved test reliability by scoping out tests that should not execute during app-only test runs
- Correctly set the lookup list for fields of type User [pschaeflein]
- Don't tokenize ~sitecollection in web parts XML [paulpascha]
- Updated base templates for March 2018 release
- Fix #1585 - Correctly handle Overwrite=false with the new pre-create of pages
- TimerJob framework reliability improvements (avoid breaking when clientcontext could not be obtained)
- Fix #1595 - Fixed provisioning issue when the AppCatalog is missing. [gautamdsheth]
- Updated DataRow handler in provisioning engine to not update readonly fields, and to allow for emptying fields by leaving the element value empty of a DataValue element.
- Support extraction of "empty" client side pages when using an extensibility provider that extracts more than the default home page
- Improved detection of illegal characters in folder and file names [aslanovsergey]
- Fix #1509 - Role inheritance can be broken when site security is specified with BreakRoleInheritance set to true without additional RoleAssignments specified [paulpascha]
- Commenting can be enabled/disabled on home page via the ClientSidePages object handler
- RoleDefinitions are now parsed in the SiteSecurity object handler
- WebHook provisioning errors will not stop the provisioning process
- Improved list content type handling [jensotto]
- Exclude ComposedLook handler processing for NoScript sites
- Improved detection of App-Only to support weblogin based use
- SiteName and SiteTitle token updates [jensotto]
- Fix #1059 - SharePoint 2013 on premise issues with ApplyProvisioningTemplate when publishing activated
- Switched to CSOM version 7414.1200
- Groupify method supports the "keep existing homepage" scenario
- Fixed behavior while adding/updating datarows with the Provisioning Engine [craig-blowfield]

### Deprecated

- Marked SetDefaultContentTypeToList extensions methods on List and Web objects as deprecated. This method has some flaws. It was possible to use the ID of a content type at site level to set as a default content type in the list, IF a content type in that list was inheriting from the parent content type. The new method requires you to specify the actual content type that is associated with the list. It will not work to specify a parent content type id.
