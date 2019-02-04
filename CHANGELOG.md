# OfficeDevPnP.Sites.Core Changelog

*Please do not commit changes to this file, it is maintained by the repo owner.*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/).

## [3.6.1902.0 - unreleased]

### Added

- Added support for creating and loading modern pages from sub folders inside the sitepages library
- ZoneID web part property now can be used in SP2016
- MajorVersionLimit and MajorWithMinorVersionsLimit are supported in the minimal (May 2018) version of SP2013 CSOM (Issue 1943) #1994 [tmeckel]
- Enables Web.RequestAccessEmail for on-premises (both 15.0 and 16.0) #1794 [biste5]
- Add token parsing in `targetFileName` property of file object #2036 [stevebeauge]
- Added support to delete search configurations
- Add support for setting default sharing and sharing permissions on tenant extensions

### Changed

- Feature/make datarow and file properties consistent #1762 [stevebeauge]

## [3.5.1901.0 - January 2019 release]

### Added

- Added support for modern page section backgrounds
- Added new 1st party client side web parts to the client side page API - support for provisioning engine will come with next schema update
- Added support for webparts configured with isDomainIsolated=true - support for provisioning engine will come with next schema update
- ResetFileToPreviousVersion extension method #2030 [skaggej]
  
### Changed

- Fix to make the EveryoneExceptExternalUsers token resolve correctly in all circumstances
- Fix to ensure TLS settings are correctly configured on certain OS versions (e.g. Windows Server 2012 R2)
- Fix throttling Retry-After processing, should be in seconds, not in milliseconds
- Multi-lingual provisioning of list title, extraction of additional navigation node languages #1974 [czullu]
- Updated logging logic #2018 [jensotto]
- Performance optimization on for the client side page save action

### Deprecated

## [3.4.1812.1 - December 2018 release]

### Added

- Added support for handling new page header options

### Changed

### Deprecated

- Deprecated Responsive UI extension methods
  
## [3.4.1812.0 - December 2018 release]

### Added

- Adding support for a 3rd navigation level in provisioning (for modern pages) #1927 [mbruckner]
- Ability to update content type properties #1776 [gautamdsheth]
- Ability to create team with Group #1990 [gautamdsheth]
- Ability to enable/disable comments, likes and view count on modern site pages #1756 [gautamdsheth]
- Added support for themes generation via ThemeUtility.GetThemeAsJSON(primaryColor, bodyTextColor, bodyBackgroundColor) [paolopia]

### Changed

- Stability improvements for updates to RoleDefinition update #1846 [sebastianmattar]
- Prevent access denied exception when provisioning content types #1903 [jensotto]
- Allow parameters in field defaults #1979 [oozoo-solutions]
- Add token parsing when provisioning search settings #1727 [jensotto]
- Fixed issue with calculated fields for non-English site collections #1970 [SchauDK]
- FixLookupField. If target list is not found, just return fieldXml #1977 [SchauDK]
- Current user can't be removed from new SecurableObject role assignments #1584 [jensotto]
- Use Xml token parsing for Xml data #1982 [SchauDK]
- New CSOM throttling implementation
- Fix: Token parser #1968 #1972 [SchauDK] [phawrylak]
- Improve add owner/member on Group creating #1987 #1990 #1991 [sadomovalex] [gautamdsheth]
- Improved handling of CustomSortOrder for terms in Term Store [TeodoraI]
- Improved Tenant and ALM handlers to avoid useless processing [gautamdsheth]

## [3.3.1811.0 - November 2018 release]

### Added

- Added support for the `Visibility` attribute for Unified Groups [devinprejean]
- Added support for language/lcid when creating modern sites using Sites.SiteCollection.CreateAsync method.
- Added support for FieldIdToken to support customers while migrating across sites and keeping field internal name, but changing field Id.
- Added support for Single Page WebPart App pages, will be part of SPFX 1.7
- Added support for Resource Path API in modern pages #1936 [gautamdsheth]

### Changed

- Get classification directly from Unified Group instead of a separate call [devinprejean]
- Removes 60 minute maximum lifetime for Access Tokens in AuthenticationManager #1957 [koskila]
- Fix: MaxVersionLimit set to 0 issue [gautamdsheth]

### Deprecated

## [3.2.1810.0 - October 2018 release]

### Added
- Added support for provisioning a site hierarchy through the provisioning engine based upon the 2018-07 schema.
- Added Tenant.ApplyProvisioningHierarchy extension method
- Added various additional provisioning engine object handlers to support sitehierarchy
- Added ability to set SiteLogo on a modern team site through Sites.SiteCollection.SetGroupImage method.

### Changed

- ClientSide page name now can contain a token [gautamdsheth]
- Fix issue with AssociatedGroupToken loading [gautamdsheth]
- LoginNames are compared case insensitive [tmeckel]
- Allow to create a CustomAction to a ListInstance without specifying a valid XML for the CommandUIExtension [tmeckel]
- Don't create a custom sort order for the HashTags TermSet [tmeckel]
- Use topological sort to order groups before creating them [tmeckel]
- Don't process web hook assignments without having a valid URL [phawrylak]
- Refactored objectterms and objectenant handler to support provisioning hierarchies.
- Don't export the internal _DisplayName field [phawrylak]
- Fixed SetOpenBySitePolicy as it never worked [gautamdsheth]
- Fixed ServerUnauthorizedAccessException when creating web (#1925) [phawrylak]

### Deprecated
- Deprecated all provisioning engine tokens that start with ~, like ~site, etc. Use {site} etc. instead. ~ tokens conflicted with a token system used by SharePoint itself.

## [3.1.1809.0 - September 2018 release]

### Added
- Added support to provision hidden views
- Added support for inviting guest users (AAD B2B) via Microsoft Graph [Vipul Kelkar]

### Changed
- Fixed issue where hidden views created by XsltListView web part where removed on a list during provisioning
- Refactored token parsing for PnP template handling for performance
- Support token replacement for view xml [vonis22]
- Updated CSOM Assemblies to 8029.1200
- Bugfix for token replacement where two tokens where next to each other like {hosturl}{siteid}
- Bugfix and optimizatin for web part listid token replacement
- Make preview link for banner image on modern pages link to the root site to avoid too long url's - and act like the default behaviour
- Fix for updating Unified Groups [Gautam Sheth]
- Extensibility handlers error handling [Jens Otto Hatlevold]
- Fix default client side page header title alignment

### Deprecated
- Marked regex functions in TokenDefinition as obsolete, as they are not needed

## [3.0.1808.0 - August 2018 release]

### Added

### Changed
- Introduced support for ADAL 3.x and JWT 5.x, updated NuGet package reference accordingly
- Client side API - Correctly handle data version: split between canvas and webpart data version + export data vesion using the provisioning engine + improved data version detection
- Bug fix for using SetDefaultColumnValues in lists in subsites [cnesmark]
- Fixed an issue with lookup fields in a list instance, when a template is applied to update a lookup field [antim-mironov]

### Deprecated

## [2.28.1807.0 - July 2018 release]

### Added
- Information management async extension methods #1843 [baywet]
- TimerJob AppOnly authentication in High Trust context #1808 [ypcode]

### Changed
- Added PowerApps client side web part type
- Fix NullReferenceException when parsing client side page header html #1821 [SchauDK]
- Changed multi lookup field provisioning to also handle list url in List #1822 [cebud]
- Don't wrap client side text in P if it already was done as part of the provided text
- Added tokenization of client side page header image url
- Fix #1810 ContentTypeBinding with lowercase ContentTypeID [TeodoraI]
- Fix list attribute for lookup fields #1826 [sebastianmattar]

### Deprecated

## [2.27.1806.0 - June 2018 release]

### Added
- Added optional timeout value on AppManager.Add method
- Support version 1.4 of page header data structure
- Feature/file folder async extension methods [baywet]

### Changed
- ClientComponentId and ClientComponentProperties are now updated when applying a template to a site where the customaction already exists [SchauDK]
- Fixes issue with requiring tenant admin access while not provisioning tenant scoped artifacts
- Fixed issue where a list would not be created based on a list template (TemplateFeatureId)
- Fixes issue with double tokens in content by search webpart provisioning [KEMiCZA]
- Fixes issue with sitedesigns not correctly being associated to web template
- Fixes issue where you could not specify content type in a datarow element in a provisioning template
- Fixes issue where you tried to modify a property of a default modern home page, and all web parts disapeared
- Fixed issue with Security Group names including HTML links [jensotto]
- Fixed issue with UseShared property for Navigation Settings [TheJeffer]
- Fixed issue with not existing links in Navigation Settings [gautamdsheth]
- Updated Microsoft Graph SDK package to version 1.9.0
- Correctly extract modern page title [SchauDK]
- Fixes issue with using culture in page header persisting [guillaume-kizilian]
- Fixes lookup column support by supporting list web relative urls [stevebeauge]
- Fixed ClientSidePageHeaderType enum inconsistency [SchauDK]
- Fixing #1770 issue. Now we are considering Publishing Images field type [luismanez]
- #1804 Incorrect exception thrown while setting multi-valued tax field [gautamdsheth]
- Typo fixes [stwel]

### Deprecated

## [2.26.1805.0 - May 2018 release]

### Added
- Added WebApiPermissions support to provisioning engine.
- Added support to auto populate the BannerImageUrl and Description fields during save of a client side page based on the found web parts and text parts on the page
- Added support for client side page header configuration (no header, header with image, default header)
- Added ClientSidePage Title support in the provisioning engine.
- Added CommentsOnSitePagesDisabled property on web settings element in the provisioning engine.
- Added support for StorageEntities to the Tenant element in the Provisioning Engine. The user applying the template needs appropriate access rights to the tenant scoped App Catalog.
- Added SiteScripts and SiteDesigns elements to the Tenant element in the Provisioning Engine. The user applying the template needs to be tenant administrator.
- Added HubSiteUrl to the WebSettings element for the Provisioning Engine. The user applying the template needs to be tenant administrator.
- Added {SiteScriptId:[script title]} and {SiteDesignId:[design title]} tokens to the provisioning engine. This will only work if the user applying the template is tenant administrator.
- Added {StorageEntityValue:[key]} token to retrieve values from tenant level or (when applicable) site collection level. If a key is present at site collection level this value will take preference over the one from tenant level, following the behavior of the CSOM APIs.
- Added support for loading the classification of a unified group.
- Added GetPrincipalUniqueRoleAssignments web extension method. Get all unique role assignments for a user or a group in a web object and all its descendents down to document or list item level.
- Added support for SystemUpdate of taxonomy fields on list extension and item extension methods.
- Added support for using the ClientWebPart client side web part to host "classic" SharePoint Add-ins on client side pages
- Added support for new schema v.2018-05
- Added support for Web API Permission in schema v.2018-05
- Added support for new schema v.2018-05 ==> 2018-05 is the new default schema
- Added async extension methods for feature handling and property retrieval [baywet]
- Added extension methods to better support property handling on lists [gautamdsheth]
- Added support for the implementation of the provisioning of dependent lookups fields [stevebeauge]

### Changed
- Fixed typo in TimeZone enum, and obsoleted incorrect value [gautamdsheth]
- Web hook server notification url in the provisioning engine now supports tokens [krzysztofziemacki]
- Fixed the setting of the page layout [TheJeffer] 
- Improved detection and configuration of the specific client side web part data version
- Allow webhooks expiration to be updated without specifying the original web hook notification url [tavikukko]
- Fixed detecting of "The object specified does not belong to a list" error in the SetFileProperties extension method [Ralmenar]
- Using ResourcePath.FromDecodedUrl to handle reading files and folders with special characters [gautamdsheth]
- Fix async handling calling ClientSidePage.AvailableClientSideComponents [OliverZeiser]

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
