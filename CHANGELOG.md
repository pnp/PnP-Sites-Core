# OfficeDevPnP.Sites.Core Changelog

*Please do not commit changes to this file, it is maintained by the repo owner.*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/).

## [2.24.1803.0 - Unreleased]

### Added

- Added EnsureLabel extension method to the taxonomy extensions [paulpascha]
- Added SetDefaultContentType extension methods on List objects. Notice that this method behaves different from the deprecated SetDefaultContentTypeToList method. See the Deprecated section.
- Added AliasExistsAsync extension method to verify if an Office 365 Group alias is available for use
- Added support for taxonomy fields in DataRows at the provisioning engine level. [jensotto]
- Added support for updating owners and members of an Office 365 Group.
- Added support for TermStore DefaultLanguage when retrieving or adding a term. [stevebeauge]
- Added support for getting apps by title [gautamdsheth]
- Added .NET 2.0 Standard project to allow cross-platform use of the PnP Sites Core library

### Changed

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
