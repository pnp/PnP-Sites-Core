# OfficeDevPnP.Core
This is automatically generate documentation for the Office 365 Dev PnP Sites Core component.
- [PnP Sites Core Component in GitHub](https://github.com/OfficeDev/PnP-Sites-Core/tree/master/Core)

## SharePoint.Client.ClientContextExtensions
            
Class that holds the deprecated clientcontext methods
        
### Methods


#### Clone(Microsoft.SharePoint.Client.ClientRuntimeContext,System.String)
Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
> ##### Parameters
> **clientContext:** ClientContext to be cloned

> **siteUrl:** Site url to be used for cloned ClientContext

> ##### Return value
> A ClientContext object created for the passed site url

#### ExecuteQueryRetry(Microsoft.SharePoint.Client.ClientRuntimeContext,System.Int32,System.Int32)

> ##### Parameters
> **clientContext:** 

> **retryCount:** Number of times to retry the request

> **delay:** Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry


#### Clone(Microsoft.SharePoint.Client.ClientRuntimeContext,System.Uri)
Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
> ##### Parameters
> **clientContext:** ClientContext to be cloned

> **siteUrl:** Site url to be used for cloned ClientContext

> ##### Return value
> A ClientContext object created for the passed site url

#### GetSiteCollectionContext(Microsoft.SharePoint.Client.ClientRuntimeContext)
Gets a site collection context for the passed web. This site collection client context uses the same credentials as the passed client context
> ##### Parameters
> **clientContext:** Client context to take the credentials from

> ##### Return value
> A site collection client context object for the site collection

#### IsAppOnly(Microsoft.SharePoint.Client.ClientRuntimeContext)
Checks if the used ClientContext is app-only
> ##### Parameters
> **clientContext:** The ClientContext to inspect

> ##### Return value
> True if app-only, false otherwise

#### Constructor
Constructor
> ##### Parameters
> **message:** 


#### HasMinimalServerLibraryVersion(Microsoft.SharePoint.Client.ClientRuntimeContext,System.String)
Checks the server library version of the context for a minimally required version
> ##### Parameters
> **clientContext:** 

> **minimallyRequiredVersion:** 

> ##### Return value
> 

## SharePoint.Client.ClientContextExtensions.MaximumRetryAttemptedException
            
Defines a Maximum Retry Attemped Exception
        
### Methods


#### Constructor
Constructor
> ##### Parameters
> **message:** 


## SharePoint.Client.BrandingExtensions
            
Class that holds the deprecated branding methods
            
Class that deals with branding features
        
### Methods


#### DisableReponsiveUI(Microsoft.SharePoint.Client.Site)
Disables the Responsive UI on a Classic SharePoint Site
> ##### Parameters
> **site:** The Site to disable the Responsive UI on


#### DisableReponsiveUI(Microsoft.SharePoint.Client.Web)
Disables the Responsive UI on a Classic SharePoint Web
> ##### Parameters
> **web:** The Web to disable the Responsive UI on


#### ComposedLookExists(Microsoft.SharePoint.Client.Web,System.String)
Checks if a composed look exists.
> ##### Parameters
> **web:** Web to check

> **composedLookName:** Name of the composed look

> ##### Return value
> true if it exists; otherwise false

#### CreateComposedLookByName(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.String,System.Int32,System.Boolean)
Creates (or updates) a composed look in the web site; usually this is done in the root site of the collection.
> ##### Parameters
> **web:** Web to create the composed look in

> **lookName:** Name of the theme

> **paletteFileName:** File name of the palette file in the theme catalog of the site collection; path component ignored.

> **fontFileName:** File name of the font file in the theme catalog of the site collection; path component ignored.

> **backgroundFileName:** File name of the background image file in the theme catalog of the site collection; path component ignored.

> **masterFileName:** File name of the master page in the mastepage catalog of the web site; path component ignored.

> **displayOrder:** Display order of the composed look

> **replaceContent:** Replace composed look if it already exists (default true)


#### CreateComposedLookByUrl(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.String,System.Int32,System.Boolean)
Creates (or updates) a composed look in the web site; usually this is done in the root site of the collection.
> ##### Parameters
> **web:** Web to create the composed look in

> **lookName:** Name of the theme

> **paletteServerRelativeUrl:** URL of the palette file, usually in the theme catalog of the site collection

> **fontServerRelativeUrl:** URL of the font file, usually in the theme catalog of the site collection

> **backgroundServerRelativeUrl:** URL of the background image file, usually in /_layouts/15/images

> **masterServerRelativeUrl:** URL of the master page, usually in the masterpage catalog of the web site

> **displayOrder:** Display order of the composed look

> **replaceContent:** Replace composed look if it already exists (default true)


#### SetComposedLookByUrl(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.String,System.Boolean,System.Boolean)
Retrieves the named composed look, overrides with specified palette, font, background and master page, and then recursively sets the specified values.
> ##### Parameters
> **web:** Web to apply composed look to

> **lookName:** Name of the composed look to apply; null will apply the override values only

> **paletteServerRelativeUrl:** Override palette file URL to use

> **fontServerRelativeUrl:** Override font file URL to use

> **backgroundServerRelativeUrl:** Override background image file URL to use

> **masterServerRelativeUrl:** Override master page file URL to use

> **resetSubsitesToInherit:** false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit

> **updateRootOnly:** false to apply to subsites; true (default) to only apply to specified site


#### SetThemeByUrl(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.Boolean,System.Boolean)
Recursively applies the specified palette, font, and background image.
> ##### Parameters
> **web:** Web to apply to

> **paletteServerRelativeUrl:** URL of palette file to apply

> **fontServerRelativeUrl:** URL of font file to apply

> **backgroundServerRelativeUrl:** URL of background image to apply

> **resetSubsitesToInherit:** false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit

> **updateRootOnly:** false (default) to apply to subsites; true to only apply to specified site


#### UploadThemeFile(Microsoft.SharePoint.Client.Web,System.String,System.String)
Uploads the specified file (usually an spcolor or spfont file) to the web site themes gallery (usually only exists in the root web of a site collection).
> ##### Parameters
> **web:** Web site to upload to

> **localFilePath:** Location of the file to be uploaded

> **themeFolderVersion:** Leaf folder name to upload to; default is "15"

> ##### Return value
> The uploaded file, with at least the ServerRelativeUrl property available

#### UploadThemeFile(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String)
Uploads the specified file (usually an spcolor or spfont file) to the web site themes gallery (usually only exists in the root web of a site collection).
> ##### Parameters
> **web:** Web site to upload to

> **fileName:** Name of the file to upload

> **localFilePath:** Location of the file to be uploaded

> **themeFolderVersion:** Leaf folder name to upload to; default is "15"

> ##### Return value
> The uploaded file, with at least the ServerRelativeUrl property available

#### UploadThemeFile(Microsoft.SharePoint.Client.Web,System.String,System.IO.Stream,System.String)
Uploads the specified file (usually an spcolor or spfont file) to the web site themes gallery (usually only exists in the root web of a site collection).
> ##### Parameters
> **web:** Web site to upload to

> **fileName:** Name of the file to upload

> **localStream:** Stream containing the contents of the file

> **themeFolderVersion:** Leaf folder name to upload to; default is "15"

> ##### Return value
> The uploaded file, with at least the ServerRelativeUrl property available

#### DeployPageLayout(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.String)
Can be used to deploy page layouts to master page gallery. Should be only used with root web of site collection where publishing features are enabled.
> ##### Parameters
> **web:** Web as the root site of the publishing site collection

> **sourceFilePath:** Full path to the file which will be uploaded

> **title:** Title for the page layout

> **description:** Description for the page layout

> **associatedContentTypeID:** Associated content type ID

> **folderHierarchy:** Folder hierarchy where the page layouts will be deployed


#### DeployHtmlPageLayout(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.String)
Can be used to deploy html page layouts to master page gallery. Should be only used with root web of site collection where publishing features are enabled.
> ##### Parameters
> **web:** Web as the root site of the publishing site collection

> **sourceFilePath:** Full path to the file which will be uploaded

> **title:** Title for the page layout

> **description:** Description for the page layout

> **associatedContentTypeID:** Associated content type ID

> **folderHierarchy:** Folder hierarchy where the html page layouts will be deployed


#### DeployMasterPageGalleryItem(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.String,System.String)
Private method to support all kinds of file uploads to the master page gallery
> ##### Parameters
> **web:** Web as the root site of the publishing site collection

> **sourceFilePath:** Full path to the file which will be uploaded

> **title:** Title for the page layout

> **description:** Description for the page layout

> **associatedContentTypeID:** Associated content type ID

> **itemContentTypeId:** Content type id for the item.

> **folderHierarchy:** Folder hierarchy where the file will be uploaded


#### DeployMasterPage(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.String,System.String)
Deploys a new masterpage
> ##### Parameters
> **web:** The web to process

> **sourceFilePath:** The path to the source file

> **title:** The title of the masterpage

> **description:** The description of the masterpage

> **uiVersion:** 

> **defaultCSSFile:** 

> **folderPath:** 


#### SetMasterPagesByName(Microsoft.SharePoint.Client.Web,System.String,System.String)
Can be used to set master page and custom master page in single command
> ##### Parameters
> **web:** 

> **masterPageName:** 

> **customMasterPageName:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when masterPageName or customMasterPageName is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when masterPageName or customMasterPageName is null


#### SetMasterPagesByUrl(Microsoft.SharePoint.Client.Web,System.String,System.String)
Can be used to set master page and custom master page in single command
> ##### Parameters
> **web:** 

> **masterPageUrl:** 

> **customMasterPageUrl:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when masterPageName or customMasterPageName is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when masterPageName or customMasterPageName is null


#### SetMasterPageByName(Microsoft.SharePoint.Client.Web,System.String)
Master page is set by using master page name. Master page is set from the current web.
> ##### Parameters
> **web:** Current web

> **masterPageName:** Name of the master page. Path is resolved from this.

> ##### Exceptions
> **System.ArgumentException:** Thrown when masterPageName is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when masterPageName is null


#### SetCustomMasterPageByName(Microsoft.SharePoint.Client.Web,System.String)
Master page is set by using master page name. Master page is set from the current web.
> ##### Parameters
> **web:** Current web

> **masterPageName:** Name of the master page. Path is resolved from this.

> ##### Exceptions
> **System.ArgumentException:** Thrown when masterPageName is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when masterPageName is null


#### GetRelativeUrlForMasterByName(Microsoft.SharePoint.Client.Web,System.String)
Returns the relative URL for a masterpage
> ##### Parameters
> **web:** 

> **masterPageName:** The name of the masterpage, e.g. 'default' or 'seattle'

> ##### Return value
> 

#### GetCurrentComposedLook(Microsoft.SharePoint.Client.Web)
Returns the current theme of a web
> ##### Parameters
> **web:** Web to check

> ##### Return value
> Entity with attributes of current composed look, or null if none

#### GetComposedLook(Microsoft.SharePoint.Client.Web,System.String)
Returns the named composed look from the web gallery
> ##### Parameters
> **web:** Web to check

> **composedLookName:** Name of the composed look to retrieve

> ##### Return value
> Entity with the attributes of the composed look, or null if the composed look does not exists or cannot be determined

#### IsMatchingTheme(OfficeDevPnP.Core.Entities.ThemeEntity,System.String,System.String,System.String)
Compares master page URL, theme URL and font URL values to current theme entity to check if they are the same. Handles also possible null values. Point is to figure out which theme is the one that is currently being selected as "Current"
> ##### Parameters
> **theme:** Current theme entity to compare values to

> **masterPageUrl:** Master page URL

> **themeUrl:** Theme URL

> **fontUrl:** Font URL

> ##### Return value
> 

#### GetPageLayoutListItemByName(Microsoft.SharePoint.Client.Web,System.String)
Gets a page layout from the master page catalog. Can be called with paramter as "pagelayout.aspx" or as full path like "_catalog/masterpage/pagelayout.aspx"
> ##### Parameters
> **web:** root web

> **pageLayoutName:** name of the page layout to retrieve

> ##### Return value
> ListItem holding the page layout, null if not found

#### SetMasterPageByUrl(Microsoft.SharePoint.Client.Web,System.String,System.Boolean,System.Boolean)
Set master page by using given URL as parameter. Suitable for example in cases where you want sub sites to reference root site master page gallery. This is typical with publishing sites.
> ##### Parameters
> **web:** Context web

> **masterPageServerRelativeUrl:** URL to the master page.

> **resetSubsitesToInherit:** false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit

> **updateRootOnly:** false (default) to apply to subsites; true to only apply to specified site


#### SetCustomMasterPageByUrl(Microsoft.SharePoint.Client.Web,System.String,System.Boolean,System.Boolean)
Set Custom master page by using given URL as parameter. Suitable for example in cases where you want sub sites to reference root site master page gallery. This is typical with publishing sites.
> ##### Parameters
> **web:** Context web

> **masterPageServerRelativeUrl:** URL to the master page.

> **resetSubsitesToInherit:** false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit

> **updateRootOnly:** false (default) to apply to subsites; true to only apply to specified site


#### SetDefaultPageLayoutForSite(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Web,System.String)
Sets specific page layout the default page layout for the particular site
> ##### Parameters
> **web:** 

> **rootWeb:** 

> **pageLayoutName:** 


#### SetSiteToInheritPageLayouts(Microsoft.SharePoint.Client.Web)
Can be used to set the site to inherit the default page layout option from parent. Cannot be used for root site of the site collection
> ##### Parameters
> **web:** Web to operate against


#### AllowAllPageLayouts(Microsoft.SharePoint.Client.Web)
Allow the web to use all available page layouts
> ##### Parameters
> **web:** Web to operate against


#### SetAvailablePageLayouts(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Web,System.Collections.Generic.IEnumerable{System.String})
Sets the available page layouts
> ##### Parameters
> **web:** The web to process

> **rootWeb:** The rootweb

> **pageLayouts:** The page layouts to make available


#### SetAvailableWebTemplates(Microsoft.SharePoint.Client.Web,System.Collections.Generic.List{OfficeDevPnP.Core.Entities.WebTemplateEntity})
Defines which templates are available for subsite creation
> ##### Parameters
> **web:** Web to operate against

> **availableTemplates:** List of objects that define the templates that are allowed


#### ClearAvailableWebTemplates(Microsoft.SharePoint.Client.Web)
Can be used to remote filters from the available web template
> ##### Parameters
> **web:** 


#### SetHomePage(Microsoft.SharePoint.Client.Web,System.String)
Sets the web home page
> ##### Parameters
> **web:** The Web to process

> **rootFolderRelativePath:** The path relative to the root folder of the site, e.g. SitePages/Home.aspx


#### EnableResponsiveUI(Microsoft.SharePoint.Client.Web,System.String)
Enables the responsive UI of a classic SharePoint Web
> ##### Parameters
> **web:** The Web to activate the Responsive UI to

> **infrastructureUrl:** URL pointing to an infrastructure site


#### EnableResponsiveUI(Microsoft.SharePoint.Client.Site,System.String)
Enables the responsive UI of a classic SharePoint Site
> ##### Parameters
> **site:** The Site to activate the Responsive UI to

> **infrastructureUrl:** URL pointing to an infrastructure site


#### EnableResponsiveUIImplementation(Microsoft.SharePoint.Client.ClientObject,System.String)
Enables the responsive UI of a classic SharePoint Web or Site
> ##### Parameters
> **clientObject:** The Web or Site to activate the Responsive UI to

> **infrastructureUrl:** URL pointing to an infrastructure site


#### DisableResponsiveUI(Microsoft.SharePoint.Client.Web)
Disables the Responsive UI on a Classic SharePoint Web
> ##### Parameters
> **web:** The Web to disable the Responsive UI on


#### DisableResponsiveUI(Microsoft.SharePoint.Client.Site)
Disables the Responsive UI on a Classic SharePoint Site
> ##### Parameters
> **site:** The Site to disable the Responsive UI on


## SharePoint.Client.FeatureExtensions
            
Class that holds deprecated feature activation and deactivation methods
            
Class that deals with feature activation and deactivation
        
### Methods


#### ActivateFeature(Microsoft.SharePoint.Client.Web,System.Guid,System.Boolean,System.Int32)
Activates a site collection or site scoped feature
> ##### Parameters
> **web:** Web to be processed - can be root web or sub web

> **featureID:** ID of the feature to activate

> **sandboxed:** Set to true if the feature is defined in a sandboxed solution

> **pollingIntervalSeconds:** The time in seconds between polls for "IsActive"


#### ActivateFeature(Microsoft.SharePoint.Client.Site,System.Guid,System.Boolean,System.Int32)
Activates a site collection or site scoped feature
> ##### Parameters
> **site:** Site to be processed

> **featureID:** ID of the feature to activate

> **sandboxed:** Set to true if the feature is defined in a sandboxed solution

> **pollingIntervalSeconds:** The time in seconds between polls for "IsActive"


#### DeactivateFeature(Microsoft.SharePoint.Client.Web,System.Guid,System.Int32)
Deactivates a site collection or site scoped feature
> ##### Parameters
> **web:** Web to be processed - can be root web or sub web

> **featureID:** ID of the feature to deactivate

> **pollingIntervalSeconds:** The time in seconds between polls for "IsActive"


#### DeactivateFeature(Microsoft.SharePoint.Client.Site,System.Guid,System.Int32)
Deactivates a site collection or site scoped feature
> ##### Parameters
> **site:** Site to be processed

> **featureID:** ID of the feature to deactivate

> **pollingIntervalSeconds:** The time in seconds between polls for "IsActive"


#### IsFeatureActive(Microsoft.SharePoint.Client.Site,System.Guid)
Checks if a feature is active
> ##### Parameters
> **site:** Site to operate against

> **featureID:** ID of the feature to check

> ##### Return value
> True if active, false otherwise

#### IsFeatureActive(Microsoft.SharePoint.Client.Web,System.Guid)
Checks if a feature is active
> ##### Parameters
> **web:** Web to operate against

> **featureID:** ID of the feature to check

> ##### Return value
> True if active, false otherwise

#### IsFeatureActiveInternal(Microsoft.SharePoint.Client.FeatureCollection,System.Guid,System.Boolean)
Checks if a feature is active in the given FeatureCollection.
> ##### Parameters
> **features:** FeatureCollection to check in

> **featureID:** ID of the feature to check

> **noRetry:** Use regular ExecuteQuery

> ##### Return value
> True if active, false otherwise

#### ProcessFeature(Microsoft.SharePoint.Client.Site,System.Guid,System.Boolean,System.Boolean,System.Int32)
Activates or deactivates a site collection scoped feature
> ##### Parameters
> **site:** Site to be processed

> **featureID:** ID of the feature to activate/deactivate

> **activate:** True to activate, false to deactivate the feature

> **sandboxed:** Set to true if the feature is defined in a sandboxed solution

> **pollingIntervalSeconds:** The time in seconds between polls for "IsActive"


#### ProcessFeature(Microsoft.SharePoint.Client.Web,System.Guid,System.Boolean,System.Boolean,System.Int32)
Activates or deactivates a web scoped feature
> ##### Parameters
> **web:** Web to be processed - can be root web or sub web

> **featureID:** ID of the feature to activate/deactivate

> **activate:** True to activate, false to deactivate the feature

> **sandboxed:** True to specify that the feature is defined in a sandboxed solution

> **pollingIntervalSeconds:** The time in seconds between polls for "IsActive"


#### ProcessFeatureInternal(Microsoft.SharePoint.Client.FeatureCollection,System.Guid,System.Boolean,Microsoft.SharePoint.Client.FeatureDefinitionScope,System.Int32)
Activates or deactivates a site collection or web scoped feature
> ##### Parameters
> **features:** Feature Collection which contains the feature

> **featureID:** ID of the feature to activate/deactivate

> **activate:** True to activate, false to deactivate the feature

> **scope:** Scope of the feature definition

> **pollingIntervalSeconds:** The time in seconds between polls for "IsActive"


## SharePoint.Client.FieldAndContentTypeExtensions
            
This class holds deprecated extension methods that will help you work with fields and content types.
            
This class provides extension methods that will help you work with fields and content types.
        
### Methods


#### CreateField(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.FieldCreationInformation,System.Boolean)
Create field to web remotely
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **fieldCreationInformation:** Creation Information for the field.

> **executeQuery:** Optionally skip the executeQuery action

> ##### Return value
> The newly created field or existing field.

#### CreateField``1(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.FieldCreationInformation,System.Boolean)
Create field to web remotely
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **fieldCreationInformation:** Field creation information

> **executeQuery:** Optionally skip the executeQuery action

> ##### Return value
> The newly created field or existing field.

#### CreateField(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Create field to web remotely
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **fieldAsXml:** The XML declaration of SiteColumn definition

> **executeQuery:** 

> ##### Return value
> The newly created field or existing field.

#### RemoveFieldByInternalName(Microsoft.SharePoint.Client.Web,System.String)
Removes a field by specifying its internal name
> ##### Parameters
> **web:** 

> **internalName:** 


#### RemoveFieldById(Microsoft.SharePoint.Client.Web,System.String)
Removes a field by specifying its ID
> ##### Parameters
> **web:** 

> **fieldId:** The id of the field to remove


#### CreateFieldsFromXMLFile(Microsoft.SharePoint.Client.Web,System.String)
Creates fields from feature element xml file schema. XML file can contain one or many field definitions created using classic feature framework structure.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **xmlFilePath:** Absolute path to the xml location


#### CreateFieldsFromXMLString(Microsoft.SharePoint.Client.Web,System.String)
Creates fields from feature element xml file schema. XML file can contain one or many field definitions created using classic feature framework structure.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **xmlStructure:** XML structure in string format


#### CreateFieldsFromXML(Microsoft.SharePoint.Client.Web,System.Xml.Linq.XDocument)
Creates field from xml structure which follows the classic feature framework structure
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **xDocument:** Actual XML document


#### FieldExistsById(Microsoft.SharePoint.Client.Web,System.Guid,System.Boolean)
Returns if the field is found
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **fieldId:** Guid for the field ID

> **searchInSiteHierarchy:** If true, search parent sites and root site

> ##### Return value
> True or false depending on the field existence

#### GetFieldById``1(Microsoft.SharePoint.Client.Web,System.Guid,System.Boolean)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **fieldId:** Guid for the field ID

> **searchInSiteHierarchy:** If true, search parent sites and root site

> ##### Return value
> Field of type TField

#### GetFieldById(Microsoft.SharePoint.Client.Web,System.Guid,System.Boolean)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **fieldId:** Guid for the field ID

> **searchInSiteHierarchy:** If true, search parent sites and root site

> ##### Return value
> Field of type TField

#### GetFieldById``1(Microsoft.SharePoint.Client.List,System.Guid)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **list:** List to be processed. Columns assoc in lists are defined on web or rootweb.

> **fieldId:** Guid for the field ID

> ##### Return value
> Field of type TField

#### GetFieldById(Microsoft.SharePoint.Client.List,System.Guid)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **list:** List to be processed. Columns assoc in lists are defined on web or rootweb.

> **fieldId:** Guid for the field ID

> ##### Return value
> Field

#### GetFieldByInternalName(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Returns the field if it exists. Null if does not exist.
> ##### Parameters
> **web:** Web to be processed

> **internalName:** If true, search parent sites and root site

> ##### Return value
> 

#### GetFieldByName``1(Microsoft.SharePoint.Client.FieldCollection,System.String)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **fields:** FieldCollection to be processed.

> **internalName:** Guid for the field ID

> ##### Return value
> Field of type TField

#### GetFieldByName(Microsoft.SharePoint.Client.FieldCollection,System.String)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **fields:** FieldCollection to be processed.

> **internalName:** Guid for the field ID

> ##### Return value
> Field

#### GetFieldByInternalName``1(Microsoft.SharePoint.Client.FieldCollection,System.String)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **fields:** FieldCollection to be processed.

> **internalName:** Internal name of the field

> ##### Return value
> Field of type TField

#### GetFieldByInternalName(Microsoft.SharePoint.Client.FieldCollection,System.String)
Returns the field if it exists. Null if it does not exist.
> ##### Parameters
> **fields:** FieldCollection to be processed.

> **internalName:** Internal name of the field

> ##### Return value
> Field

#### FieldExistsByName(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Returns if the field is found
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **fieldName:** String for the field internal name to be used as query criteria

> **searchInSiteHierarchy:** If true, search parent sites and root site

> ##### Return value
> True or false depending on the field existence

#### FieldExistsById(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Does field exist in web
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **fieldId:** String representation of the field ID (=guid)

> **searchInSiteHierarchy:** If true, search parent sites and root site

> ##### Return value
> True if exists, false otherwise

#### FieldExistsByNameInContentType(Microsoft.SharePoint.Client.Web,System.String,System.String)
Field exists in content type
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site. Site columns should be created to root site.

> **contentTypeName:** Name of the content type

> **fieldName:** Name of the field

> ##### Return value
> True if exists, false otherwise

#### FieldExistsByNameInContentType(Microsoft.SharePoint.Client.ContentType,System.String)
Checks if a field exists in a content type by id
> ##### Parameters
> **contentType:** The content type to check

> **fieldName:** The name of the field to look for

> ##### Return value
> True if field exists in content type, otherwise false

#### SetJsLinkCustomizations(Microsoft.SharePoint.Client.Field,System.String)
Adds jsLink to a field.
> ##### Parameters
> **field:** The field to add jsLink to

> **jsLink:** JSLink to set to the form. Set to empty string to remove the set JSLink customization. Specify multiple values separated by pipe symbol. For e.g.: ~sitecollection/_catalogs/masterpage/jquery-2.1.0.min.js|~sitecollection/_catalogs/masterpage/custom.js


#### CreateField(Microsoft.SharePoint.Client.List,OfficeDevPnP.Core.Entities.FieldCreationInformation,System.Boolean)
Adds field to a list
> ##### Parameters
> **list:** List to process

> **fieldCreationInformation:** Creation information for the field

> **executeQuery:** 

> ##### Return value
> The newly created field or existing field.

#### CreateField``1(Microsoft.SharePoint.Client.List,OfficeDevPnP.Core.Entities.FieldCreationInformation,System.Boolean)
Adds field to a list
> ##### Parameters
> **list:** List to process

> **fieldCreationInformation:** Field creation information

> **executeQuery:** Optionally skip the executeQuery action

> ##### Return value
> The newly created field or existing field.

#### CreateFieldBase``1(Microsoft.SharePoint.Client.FieldCollection,OfficeDevPnP.Core.Entities.FieldCreationInformation,System.Boolean)
Base implementation for creating fields
> ##### Parameters
> **fields:** Field collection to which the created field will be added

> **fieldCreationInformation:** The information about the field to be created

> **executeQuery:** Optionally skip the executeQuery action

> ##### Return value
> 

#### FormatFieldXml(OfficeDevPnP.Core.Entities.FieldCreationInformation)
Formats a fieldcreationinformation object into Field CAML xml.
> ##### Parameters
> **fieldCreationInformation:** 

> ##### Return value
> 

#### CreateField(Microsoft.SharePoint.Client.List,System.String,System.Boolean)
Adds a field to a list
> ##### Parameters
> **list:** List to process

> **fieldAsXml:** The XML declaration of SiteColumn definition

> **executeQuery:** Optionally skip the executeQuery action

> ##### Return value
> The newly created field or existing field.

#### FieldExistsById(Microsoft.SharePoint.Client.List,System.Guid)
Returns if the field is found
> ##### Parameters
> **list:** List to process

> **fieldId:** Guid of the field ID

> ##### Return value
> True if the fields exists, false otherwise

#### FieldExistsById(Microsoft.SharePoint.Client.List,System.String)
Returns if the field is found, query based on the ID
> ##### Parameters
> **list:** List to process

> **fieldId:** String representation of the field ID (=guid)

> ##### Return value
> True if the fields exists, false otherwise

#### FieldExistsByName(Microsoft.SharePoint.Client.List,System.String)
Field exists in list by name
> ##### Parameters
> **list:** List to process

> **fieldName:** Internal name of the field

> ##### Return value
> True if the fields exists, false otherwise

#### GetFields(Microsoft.SharePoint.Client.List,System.String[])
Gets a list of fields from a list by names.
> ##### Parameters
> **list:** The target list containing the fields.

> **fieldInternalNames:** List of field names to retreieve.

> ##### Return value
> List of fields requested.

#### SetJsLinkCustomizations(Microsoft.SharePoint.Client.List,System.String,System.String)
Adds jsLink to a list field.
> ##### Parameters
> **list:** The list where the field exists.

> **fieldName:** The field to add jsLink to.

> **jsLink:** JSLink to set to the form. Set to empty string to remove the set JSLink customization. Specify multiple values separated by pipe symbol. For e.g.: ~sitecollection/_catalogs/masterpage/jquery-2.1.0.min.js|~sitecollection/_catalogs/masterpage/custom.js


#### ParseAdditionalAttributes(System.String)
Helper method to parse Key="Value" strings into a keyvaluepair
> ##### Parameters
> **xmlAttributes:** 

> ##### Return value
> 

#### AddContentTypeToListById(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Boolean,System.Boolean)
Adds content type to list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list

> **contentTypeId:** Complete ID for the content type

> **defaultContent:** Optionally make this the default content type

> **searchContentTypeInSiteHierarchy:** search for content type in site hierarchy


#### AddContentTypeToListByName(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Boolean,System.Boolean)
Adds content type to list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list

> **contentTypeName:** Name of the content type

> **defaultContent:** Optionally make this the default content type

> **searchContentTypeInSiteHierarchy:** search for content type in site hierarchy


#### AddContentTypeToList(Microsoft.SharePoint.Client.Web,System.String,Microsoft.SharePoint.Client.ContentType,System.Boolean)
Adds content type to list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list

> **contentType:** Content type to be added to the list

> **defaultContent:** If set true, content type is updated to be default content type for the list


#### AddContentTypeToListById(Microsoft.SharePoint.Client.List,System.String,System.Boolean,System.Boolean)
Add content type to list
> ##### Parameters
> **list:** List to add content type to

> **contentTypeID:** Complete ID for the content type

> **defaultContent:** If set true, content type is updated to be default content type for the list

> **searchContentTypeInSiteHierarchy:** search for content type in site hierarchy


#### AddContentTypeToListByName(Microsoft.SharePoint.Client.List,System.String,System.Boolean,System.Boolean)
Add content type to list
> ##### Parameters
> **list:** List to add content type to

> **contentTypeName:** Name of the content type

> **defaultContent:** If set true, content type is updated to be default content type for the list

> **searchContentTypeInSiteHierarchy:** search for content type in site hierarchy


#### AddContentTypeToList(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.ContentType,System.Boolean)
Add content type to list
> ##### Parameters
> **list:** List to add content type to

> **contentType:** Content type to add to the list

> **defaultContent:** If set true, content type is updated to be default content type for the list


#### AddFieldById(Microsoft.SharePoint.Client.ContentType,System.String,System.Boolean,System.Boolean)
Associates field to content type
> ##### Parameters
> **contentType:** Content Type to add the field to

> **fieldId:** String representation of the id of the field (=Guid)

> **required:** True if the field is required

> **hidden:** True if the field is hidden


#### AddFieldById(Microsoft.SharePoint.Client.ContentType,System.Guid,System.Boolean,System.Boolean)
Associates field to content type
> ##### Parameters
> **contentType:** Content Type to add the field to

> **fieldId:** The Id of the field

> **required:** True if the field is required

> **hidden:** True if the field is hidden


#### AddFieldByName(Microsoft.SharePoint.Client.ContentType,System.String,System.Boolean,System.Boolean)
Associates field to content type
> ##### Parameters
> **contentType:** Content Type to add the field to

> **fieldName:** The title or internal name of the field

> **required:** True if the field is required

> **hidden:** True if the field is hidden


#### AddFieldToContentTypeById(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Boolean,System.Boolean)
Associates field to content type
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **contentTypeID:** String representation of the id of the content type to add the field to

> **fieldId:** String representation of the field ID (=guid)

> **required:** True if the field is required

> **hidden:** True if the field is hidden


#### AddFieldToContentTypeByName(Microsoft.SharePoint.Client.Web,System.String,System.Guid,System.Boolean,System.Boolean)
Associates field to content type
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **contentTypeName:** Name of the content type

> **fieldID:** Guid representation of the field ID

> **required:** True if the field is required

> **hidden:** True if the field is hidden


#### AddFieldToContentType(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.ContentType,Microsoft.SharePoint.Client.Field,System.Boolean,System.Boolean)
Associates field to content type
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **contentType:** Content type to associate field to

> **field:** Field to associate to the content type

> **required:** Optionally make this a required field

> **hidden:** Optionally make this a hidden field


#### BestMatchContentTypeId(Microsoft.SharePoint.Client.List,System.String)
If the search finds multiple matches, the shorter ID is returned. For example, if 0x0101 is the argument, and the collection contains both 0x010109 and 0x01010901, the method returns 0x010109.
Searches the list content types and returns the content type identifier (ID) that is the nearest match to the specified content type ID.
> ##### Parameters
> **list:** The list to check for content types

> **baseContentTypeId:** A string with the base content type ID to match.

> ##### Return value
> The value of the Id property for the content type with the closest match to the value of the specified content type ID.

#### ContentTypeExistsById(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Does content type exists in the web
> ##### Parameters
> **web:** Web to be processed

> **contentTypeId:** Complete ID for the content type

> **searchInSiteHierarchy:** Searches accross all content types in the site up to the root site

> ##### Return value
> True if the content type exists, false otherwise

#### ContentTypeExistsByName(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Does content type exists in the web
> ##### Parameters
> **web:** Web to be processed

> **contentTypeName:** Name of the content type

> **searchInSiteHierarchy:** Searches accross all content types in the site up to the root site

> ##### Return value
> True if the content type exists, false otherwise

#### ContentTypeExistsById(Microsoft.SharePoint.Client.Web,System.String,System.String)
Does content type exist in web
> ##### Parameters
> **web:** Web to be processed

> **listTitle:** Title of the list to be updated

> **contentTypeId:** Complete ID for the content type

> ##### Return value
> True if the content type exists, false otherwise

#### ContentTypeExistsById(Microsoft.SharePoint.Client.List,System.String)
Does content type exist in list
> ##### Parameters
> **list:** List to update

> **contentTypeId:** Complete ID for the content type

> ##### Return value
> True if the content type exists, false otherwise

#### ContentTypeExistsByName(Microsoft.SharePoint.Client.Web,System.String,System.String)
Does content type exist in web
> ##### Parameters
> **web:** Web to be processed

> **listTitle:** Title of the list to be updated

> **contentTypeName:** Name of the content type

> ##### Return value
> True if the content type exists, false otherwise

#### ContentTypeExistsByName(Microsoft.SharePoint.Client.List,System.String)
Does content type exist in list
> ##### Parameters
> **list:** List to update

> **contentTypeName:** Name of the content type

> ##### Return value
> True if the content type exists, false otherwise

#### CreateContentTypeFromXMLFile(Microsoft.SharePoint.Client.Web,System.String)
Create a content type based on the classic feature framework structure.
> ##### Parameters
> **web:** Web to operate against

> **absolutePathToFile:** Absolute path to the xml location


#### CreateContentTypeFromXMLString(Microsoft.SharePoint.Client.Web,System.String)
Create a content type based on the classic feature framework structure.
> ##### Parameters
> **web:** Web to operate against

> **xmlStructure:** XML structure in string format


#### CreateContentTypeFromXML(Microsoft.SharePoint.Client.Web,System.Xml.Linq.XDocument)
Create a content type based on the classic feature framework structure.
> ##### Parameters
> **web:** Web to operate against

> **xDocument:** Actual XML document


#### CreateContentType(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String)
Create new content type to web
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **name:** Name of the content type

> **id:** Complete ID for the content type

> **group:** Group for the content type

> ##### Return value
> 

#### CreateContentType(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,Microsoft.SharePoint.Client.ContentType)
Create new content type to web
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **name:** Name of the content type

> **description:** Description for the content type

> **id:** Complete ID for the content type

> **group:** Group for the content type

> **parentContentType:** Parent Content Type

> ##### Return value
> The created content type

#### DeleteContentTypeByName(Microsoft.SharePoint.Client.Web,System.String)
Deletes a content type from the web by name
> ##### Parameters
> **web:** Web to delete the content type from

> **contentTypeName:** Name of the content type to delete


#### DeleteContentTypeById(Microsoft.SharePoint.Client.Web,System.String)
Deletes a content type from the web by id
> ##### Parameters
> **web:** Web to delete the content type from

> **contentTypeId:** Id of the content type to delete


#### GetContentTypeByName(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Return content type by name
> ##### Parameters
> **web:** Web to be processed

> **contentTypeName:** Name of the content type

> **searchInSiteHierarchy:** Searches accross all content types in the site up to the root site

> ##### Return value
> Content type object or null if was not found

#### GetContentTypeById(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Return content type by Id
> ##### Parameters
> **web:** Web to be processed

> **contentTypeId:** Complete ID for the content type

> **searchInSiteHierarchy:** Searches accross all content types in the site up to the root site

> ##### Return value
> Content type object or null if was not found

#### GetContentTypeByName(Microsoft.SharePoint.Client.List,System.String)
Return content type by name
> ##### Parameters
> **list:** List to update

> **contentTypeName:** Name of the content type

> ##### Return value
> Content type object or null if was not found

#### GetContentTypeById(Microsoft.SharePoint.Client.List,System.String)
Return content type by Id
> ##### Parameters
> **list:** List to update

> **contentTypeId:** Complete ID for the content type

> ##### Return value
> Content type object or null if was not found

#### BestMatch(Microsoft.SharePoint.Client.ContentTypeCollection,System.String)
Searches for the content type with the closest match to the value of the specified content type ID. If the search finds two matches, the shorter ID is returned.
> ##### Parameters
> **contentTypes:** Content type collection to search

> **contentTypeId:** Complete ID for the content type to search

> ##### Return value
> Content type Id object or null if was not found

#### RemoveContentTypeFromListByName(Microsoft.SharePoint.Client.Web,System.String,System.String)
Removes content type from list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list

> **contentTypeName:** The name of the content type


#### RemoveContentTypeFromListByName(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.List,System.String)
Removes content type from list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **list:** The List

> **contentTypeName:** The name of the content type


#### RemoveContentTypeFromListById(Microsoft.SharePoint.Client.Web,System.String,System.String)
Removes content type from a list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list

> **contentTypeId:** Complete ID for the content type


#### RemoveContentTypeFromListById(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.List,System.String)
Removes content type from a list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **list:** The List

> **contentTypeId:** Complete ID for the content type


#### RemoveContentTypeFromList(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.ContentType)
Removes content type from a list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **list:** The List

> **contentType:** The Content Type


#### SetDefaultContentTypeToList(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.List,System.String)
Set default content type to list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **list:** List to update

> **contentTypeId:** Complete ID for the content type


#### SetDefaultContentTypeToList(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.ContentType)
Set default content type to list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **list:** List to update

> **contentType:** Content type to make default


#### SetDefaultContentTypeToList(Microsoft.SharePoint.Client.Web,System.String,System.String)
Set default content type to list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list to be updated

> **contentTypeId:** Complete ID for the content type


#### SetDefaultContentTypeToList(Microsoft.SharePoint.Client.Web,System.String,Microsoft.SharePoint.Client.ContentType)
Notice. Currently removes other content types from the list. Known issue
Set's default content type list.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list to be updated

> **contentType:** Content type to make default


#### SetDefaultContentTypeToList(Microsoft.SharePoint.Client.List,System.String)
Notice. Currently removes other content types from the list. Known issue
Set's default content type list.
> ##### Parameters
> **list:** List to update

> **contentTypeId:** Complete ID for the content type


#### SetDefaultContentTypeToList(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.ContentType)
Set default content type to list
> ##### Parameters
> **list:** List to update

> **contentType:** Content type to make default


#### ReorderContentTypes(Microsoft.SharePoint.Client.List,System.Collections.Generic.IEnumerable{System.String})
Reorders content types on the list. The first one in the list is the default item. Any items left out from the list will still be on the content type, but will not be visible on the new button.
> ##### Parameters
> **list:** Target list containing the content types

> **contentTypeNamesOrIds:** Content type names or ids to sort.


#### SetLocalizationForContentType(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String)
Set localized labels for content type
> ##### Parameters
> **web:** Web to operate on

> **contentTypeName:** Name of the content type

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **nameResource:** Localized value for the Name property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForContentType(Microsoft.SharePoint.Client.List,System.String,System.String,System.String,System.String)
Set localized labels for content type
> ##### Parameters
> **list:** List to update

> **contentTypeId:** Complete ID for the content type

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **nameResource:** Localized value for the Name property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForContentType(Microsoft.SharePoint.Client.ContentType,System.String,System.String,System.String)
Set localized labels for content type
> ##### Parameters
> **contentType:** Name of the content type

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **nameResource:** Localized value for the Name property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForField(Microsoft.SharePoint.Client.Web,System.Guid,System.String,System.String,System.String)
Set localized labels for field
> ##### Parameters
> **web:** Web to operate on

> **siteColumnId:** Guid with the site column ID

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **titleResource:** Localized value for the Title property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForField(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String)
Set localized labels for field
> ##### Parameters
> **web:** Web to operate on

> **siteColumnName:** Name of the site column

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **titleResource:** Localized value for the Title property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForField(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Field,System.String,System.String,System.String)
Set localized labels for field
> ##### Parameters
> **web:** Web to operate on

> **siteColumn:** Site column to localize

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **titleResource:** Localized value for the Title property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForField(Microsoft.SharePoint.Client.List,System.Guid,System.String,System.String,System.String)
Set localized labels for field
> ##### Parameters
> **list:** List to update

> **siteColumnId:** Guid of the site column ID

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **titleResource:** Localized value for the Title property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForField(Microsoft.SharePoint.Client.List,System.String,System.String,System.String,System.String)
Set localized labels for field
> ##### Parameters
> **list:** List to update

> **siteColumnName:** Name of the site column

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **titleResource:** Localized value for the Title property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForField(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.Field,System.String,System.String,System.String)
Set localized labels for field
> ##### Parameters
> **list:** List to update

> **siteColumn:** Site column to update

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **titleResource:** Localized value for the Title property

> **descriptionResource:** Localized value for the Description property


#### SetLocalizationForField(Microsoft.SharePoint.Client.Field,System.String,System.String,System.String)
Set localized labels for field
> ##### Parameters
> **field:** Field to update

> **cultureName:** Culture for the localization (en-es, nl-be, fi-fi,...)

> **titleResource:** Localized value for the Title property

> **descriptionResource:** Localized value for the Description property


## SharePoint.Client.FileFolderExtensions
            
Class that holds the deprecated file and folder methods
        
### Methods


#### ApproveFile(Microsoft.SharePoint.Client.Web,System.String,System.String)
Approves a file
> ##### Parameters
> **web:** The web to process

> **serverRelativeUrl:** The server relative url of the file to approve

> **comment:** Message to be recorded with the approval


#### CheckInFile(Microsoft.SharePoint.Client.Web,System.String,Microsoft.SharePoint.Client.CheckinType,System.String)
Checks in a file
> ##### Parameters
> **web:** The web to process

> **serverRelativeUrl:** The server relative url of the file to checkin

> **checkinType:** The type of the checkin

> **comment:** Message to be recorded with the approval


#### CheckOutFile(Microsoft.SharePoint.Client.Web,System.String)
Checks out a file
> ##### Parameters
> **web:** The web to process

> **serverRelativeUrl:** The server relative url of the file to checkout


#### CreateDocumentSet(Microsoft.SharePoint.Client.Folder,System.String,Microsoft.SharePoint.Client.ContentTypeId)
var setContentType = list.BestMatchContentTypeId(BuiltInContentTypeId.DocumentSet); var set1 = list.RootFolder.CreateDocumentSet("Set 1", setContentType);
Creates a new document set as a child of an existing folder, with the specified content type ID.
> ##### Parameters
> **folder:** 

> **documentSetName:** 

> **contentTypeId:** Content type of the document set

> ##### Return value
> The created Folder representing the document set, so that additional operations (such as setting properties) can be done.

#### ConvertFolderToDocumentSet(Microsoft.SharePoint.Client.List,System.String)
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
Converts a folder with the given name as a child of the List RootFolder.
> ##### Parameters
> **list:** List in which the folder exists

> **folderName:** Folder name to convert

> ##### Return value
> The newly converted Document Set, so that additional operations (such as setting properties) can be done.

#### ConvertFolderToDocumentSet(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.Folder)
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
Converts a folder with the given name as a child of the List RootFolder.
> ##### Parameters
> **list:** List in which the folder exists

> **folder:** Folder to convert

> ##### Return value
> The newly converted Document Set, so that additional operations (such as setting properties) can be done.

#### ConvertFolderToDocumentSetImplementation(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.Folder)
Internal implementation of the Folder conversion to Document set
> ##### Parameters
> **list:** Library in which the folder exists

> **folder:** Folder to convert

> ##### Return value
> The newly converted Document Set, so that additional operations (such as setting properties) can be done.

#### CreateFolder(Microsoft.SharePoint.Client.Web,System.String)
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
Creates a folder with the given name as a child of the Web. Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
> ##### Parameters
> **web:** Web to check for the named folder

> **folderName:** Folder name to retrieve or create

> ##### Return value
> The newly created Folder, so that additional operations (such as setting properties) can be done.

#### CreateFolder(Microsoft.SharePoint.Client.Folder,System.String)
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters. var folder = list.RootFolder.CreateFolder("new-folder");
Creates a folder with the given name.
> ##### Parameters
> **parentFolder:** Parent folder to create under

> **folderName:** Folder name to retrieve or create

> ##### Return value
> The newly created folder

#### DoesFolderExists(Microsoft.SharePoint.Client.Web,System.String)
Checks if a specific folder exists
> ##### Parameters
> **web:** The web to process

> **serverRelativeFolderUrl:** Folder to check

> ##### Return value
> 

#### EnsureFolder(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Folder,System.String,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.Folder,System.Object}}[])
Ensure that the folder structure is created. This also ensures hierarchy of folders.
> ##### Parameters
> **web:** Web to be processed - can be root web or sub site

> **parentFolder:** Parent folder

> **folderPath:** Folder path

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> The folder structure

#### EnsureFolder(Microsoft.SharePoint.Client.Web,System.String,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.Folder,System.Object}}[])
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
Checks if the folder exists at the top level of the web site, and if it does not exist creates it. Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
> ##### Parameters
> **web:** Web to check for the named folder

> **folderName:** Folder name to retrieve or create

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> The existing or newly created folder

#### EnsureFolder(Microsoft.SharePoint.Client.Folder,System.String,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.Folder,System.Object}}[])
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
Checks if the subfolder exists, and if it does not exist creates it.
> ##### Parameters
> **parentFolder:** Parent folder to create under

> **folderName:** Folder name to retrieve or create

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> The existing or newly created folder

#### EnsureFolderPath(Microsoft.SharePoint.Client.Web,System.String,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.Folder,System.Object}}[])
If the specified path is inside an existing list, then the folder is created inside that list. Any existing folders are traversed, and then any remaining parts of the path are created as new folders.
Check if a folder exists with the specified path (relative to the web), and if not creates it (inside a list if necessary)
> ##### Parameters
> **web:** Web to check for the specified folder

> **webRelativeUrl:** Path to the folder, relative to the web site

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> The existing or newly created folder

#### FindFiles(Microsoft.SharePoint.Client.Web,System.String)
Finds files in the web. Can be slow.
> ##### Parameters
> **web:** The web to process

> **match:** a wildcard pattern to match

> ##### Return value
> A list with the found objects

#### FindFiles(Microsoft.SharePoint.Client.List,System.String)
Find files in the list, Can be slow.
> ##### Parameters
> **list:** The list to process

> **match:** a wildcard pattern to match

> ##### Return value
> A list with the found objects

#### FindFiles(Microsoft.SharePoint.Client.Folder,System.String)
Find files in a specific folder
> ##### Parameters
> **folder:** The folder to process

> **match:** a wildcard pattern to match

> ##### Return value
> A list with the found objects

#### FolderExists(Microsoft.SharePoint.Client.Web,System.String)
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
Checks if the folder exists at the top level of the web site.
> ##### Parameters
> **web:** Web to check for the named folder

> **folderName:** Folder name to retrieve

> ##### Return value
> true if the folder exists; false otherwise

#### FolderExists(Microsoft.SharePoint.Client.Folder,System.String)
Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
Checks if the subfolder exists.
> ##### Parameters
> **parentFolder:** Parent folder to check for the named subfolder

> **folderName:** Folder name to retrieve

> ##### Return value
> true if the folder exists; false otherwise

#### GetFileAsString(Microsoft.SharePoint.Client.Web,System.String)
Returns a file as string
> ##### Parameters
> **web:** The Web to process

> **serverRelativeUrl:** The server relative url to the file

> ##### Return value
> The file contents as a string

#### PublishFile(Microsoft.SharePoint.Client.Web,System.String,System.String)
Publishes a file existing on a server url
> ##### Parameters
> **web:** The web to process

> **serverRelativeUrl:** the server relative url of the file to publish

> **comment:** Comment recorded with the publish action


#### ResolveSubFolder(Microsoft.SharePoint.Client.Folder,System.String)
Gets a folder with a given name in a given
> ##### Parameters
> **folder:** in which to search for

> **folderName:** Name of the folder to search for

> ##### Return value
> The found if available, null otherwise

#### SaveFileToLocal(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.Func{System.String,System.Boolean})
Saves a remote file to a local folder
> ##### Parameters
> **web:** The Web to process

> **serverRelativeUrl:** The server relative url to the file

> **localPath:** The local folder

> **localFileName:** The local filename. If null the filename of the file on the server will be used

> **fileExistsCallBack:** Optional callback function allowing to provide feedback if the file should be overwritten if it exists. The function requests a bool as return value and the string input contains the name of the file that exists.


#### UploadFile(Microsoft.SharePoint.Client.Folder,System.String,System.String,System.Boolean)
Uploads a file to the specified folder.
> ##### Parameters
> **folder:** Folder to upload file to.

> **fileName:** 

> **localFilePath:** Location of the file to be uploaded.

> **overwriteIfExists:** true (default) to overwite existing files

> ##### Return value
> The uploaded File, so that additional operations (such as setting properties) can be done.

#### UploadFile(Microsoft.SharePoint.Client.Folder,System.String,System.IO.Stream,System.Boolean)
Uploads a file to the specified folder.
> ##### Parameters
> **folder:** Folder to upload file to.

> **fileName:** Location of the file to be uploaded.

> **stream:** 

> **overwriteIfExists:** true (default) to overwite existing files

> ##### Return value
> The uploaded File, so that additional operations (such as setting properties) can be done.

#### UploadFileWebDav(Microsoft.SharePoint.Client.Folder,System.String,System.String,System.Boolean)
Uploads a file to the specified folder by saving the binary directly (via webdav).
> ##### Parameters
> **folder:** Folder to upload file to.

> **fileName:** 

> **localFilePath:** Location of the file to be uploaded.

> **overwriteIfExists:** true (default) to overwite existing files

> ##### Return value
> The uploaded File, so that additional operations (such as setting properties) can be done.

#### UploadFileWebDav(Microsoft.SharePoint.Client.Folder,System.String,System.IO.Stream,System.Boolean)
Uploads a file to the specified folder by saving the binary directly (via webdav). Note: this method does not work using app only token.
> ##### Parameters
> **folder:** Folder to upload file to.

> **fileName:** Location of the file to be uploaded.

> **stream:** 

> **overwriteIfExists:** true (default) to overwite existing files

> ##### Return value
> The uploaded File, so that additional operations (such as setting properties) can be done.

#### GetFile(Microsoft.SharePoint.Client.Folder,System.String)
Gets a file in a document library.
> ##### Parameters
> **folder:** Folder containing the target file.

> **fileName:** File name.

> ##### Return value
> The target file if found, null if no file is found.

#### VerifyIfUploadRequired(Microsoft.SharePoint.Client.File,System.String)
Used to compare the server file to the local file. This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
> ##### Parameters
> **serverFile:** File located on the server.

> **localFile:** File to validate against.


#### VerifyIfUploadRequired(Microsoft.SharePoint.Client.File,System.IO.Stream)
Used to compare the server file to the local file. This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
> ##### Parameters
> **serverFile:** File located on the server.

> **localStream:** Stream to validate against.

> ##### Return value
> 

#### SetFileProperties(Microsoft.SharePoint.Client.File,System.Collections.Generic.IDictionary{System.String,System.String},System.Boolean)
Sets file properties using a dictionary.
> ##### Parameters
> **file:** Target file object.

> **properties:** Dictionary of properties to set.

> **checkoutIfRequired:** Check out the file if necessary to set properties.


#### PublishFileToLevel(Microsoft.SharePoint.Client.File,Microsoft.SharePoint.Client.FileLevel)
Publishes a file based on the type of versioning required on the parent library.
> ##### Parameters
> **file:** Target file to publish.

> **level:** Target publish direction (Draft and Published only apply, Checkout is ignored).


## SharePoint.Client.InformationManagementExtensions
            
Class that holds deprecated information management extension methods
            
Class that deals with information management features
        
### Methods


#### HasSitePolicyApplied(Microsoft.SharePoint.Client.Web)
Does this web have a site policy applied?
> ##### Parameters
> **web:** Web to operate on

> ##### Return value
> True if a policy has been applied, false otherwise

#### GetSiteExpirationDate(Microsoft.SharePoint.Client.Web)
Gets the site expiration date
> ##### Parameters
> **web:** Web to operate on

> ##### Return value
> DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied

#### GetSiteCloseDate(Microsoft.SharePoint.Client.Web)
Gets the site closure date
> ##### Parameters
> **web:** Web to operate on

> ##### Return value
> DateTime value holding the closure date, DateTime.MinValue in case there was no policy applied

#### GetSitePolicies(Microsoft.SharePoint.Client.Web)
Gets a list of the available site policies
> ##### Parameters
> **web:** Web to operate on

> ##### Return value
> A list of objects

#### GetAppliedSitePolicy(Microsoft.SharePoint.Client.Web)
Gets the site policy that currently is applied
> ##### Parameters
> **web:** Web to operate on

> ##### Return value
> A object holding the applied policy

#### GetSitePolicyByName(Microsoft.SharePoint.Client.Web,System.String)
Gets the site policy with the given name
> ##### Parameters
> **web:** Web to operate on

> **sitePolicy:** Site policy to fetch

> ##### Return value
> A object holding the fetched policy

#### ApplySitePolicy(Microsoft.SharePoint.Client.Web,System.String)
Apply a policy to a site
> ##### Parameters
> **web:** Web to operate on

> **sitePolicy:** Policy to apply

> ##### Return value
> True if applied, false otherwise

#### IsClosedBySitePolicy(Microsoft.SharePoint.Client.Web)
Check if a site is closed
> ##### Parameters
> **web:** Web to operate on

> ##### Return value
> True if site is closed, false otherwise

#### SetClosedBySitePolicy(Microsoft.SharePoint.Client.Web)
Close a site, if it has a site policy applied and is currently not closed
> ##### Parameters
> **web:** 

> ##### Return value
> True if site was closed, false otherwise

#### SetOpenBySitePolicy(Microsoft.SharePoint.Client.Web)
Open a site, if it has a site policy applied and is currently closed
> ##### Parameters
> **web:** 

> ##### Return value
> True if site was opened, false otherwise

## SharePoint.Client.JavaScriptExtensions
            
Deprecated JavaScript related methods
            
JavaScript related methods
        
### Fields

#### SCRIPT_LOCATION
Default Script Location value
### Methods


#### AddJsLink(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Int32)
Injects links to javascript files via a adding a custom action to the site
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **key:** Identifier (key) for the custom action that will be created

> **scriptLinks:** semi colon delimited list of links to javascript files

> **sequence:** 

> ##### Return value
> True if action was ok

#### AddJsLink(Microsoft.SharePoint.Client.Site,System.String,System.String,System.Int32)
Injects links to javascript files via a adding a custom action to the site
> ##### Parameters
> **site:** Site to be processed

> **key:** Identifier (key) for the custom action that will be created

> **scriptLinks:** semi colon delimited list of links to javascript files

> **sequence:** 

> ##### Return value
> True if action was ok

#### AddJsLink(Microsoft.SharePoint.Client.Web,System.String,System.Collections.Generic.IEnumerable{System.String},System.Int32)
Injects links to javascript files via a adding a custom action to the site
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **key:** Identifier (key) for the custom action that will be created

> **scriptLinks:** IEnumerable list of links to javascript files

> **sequence:** 

> ##### Return value
> True if action was ok

#### AddJsLink(Microsoft.SharePoint.Client.Site,System.String,System.Collections.Generic.IEnumerable{System.String},System.Int32)
Injects links to javascript files via a adding a custom action to the site
> ##### Parameters
> **site:** Site to be processed

> **key:** Identifier (key) for the custom action that will be created

> **scriptLinks:** IEnumerable list of links to javascript files

> **sequence:** 

> ##### Return value
> True if action was ok

#### DeleteJsLink(Microsoft.SharePoint.Client.Web,System.String)
Removes the custom action that triggers the execution of a javascript link
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **key:** Identifier (key) for the custom action that will be deleted

> ##### Return value
> True if action was ok

#### DeleteJsLink(Microsoft.SharePoint.Client.Site,System.String)
Removes the custom action that triggers the execution of a javascript link
> ##### Parameters
> **site:** Site to be processed

> **key:** Identifier (key) for the custom action that will be deleted

> ##### Return value
> True if action was ok

#### AddJsBlock(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Int32)
Injects javascript via a adding a custom action to the site
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **key:** Identifier (key) for the custom action that will be created

> **scriptBlock:** Javascript to be injected

> **sequence:** 

> ##### Return value
> True if action was ok

#### AddJsBlock(Microsoft.SharePoint.Client.Site,System.String,System.String,System.Int32)
Injects javascript via a adding a custom action to the site
> ##### Parameters
> **site:** Site to be processed

> **key:** Identifier (key) for the custom action that will be created

> **scriptBlock:** Javascript to be injected

> **sequence:** 

> ##### Return value
> True if action was ok

#### ExistsJsLink(Microsoft.SharePoint.Client.Web,System.String)
Checks if the target web already has a custom JsLink with a specified key
> ##### Parameters
> **web:** Web to be processed

> **key:** Identifier (key) for the custom action that will be created

> ##### Return value
> 

#### ExistsJsLink(Microsoft.SharePoint.Client.Site,System.String)
Checks if the target site already has a custom JsLink with a specified key
> ##### Parameters
> **site:** Site to be processed

> **key:** Identifier (key) for the custom action that will be created

> ##### Return value
> 

## SharePoint.Client.ListExtensions
            
Class that holds deprecated generic list creation and manipulation methods
            
Class that provides generic list creation and manipulation methods
        
### Fields

#### UrlDelimiters
The common URL delimiters
### Methods


#### AddRemoteEventReceiver(Microsoft.SharePoint.Client.List,System.String,System.String,Microsoft.SharePoint.Client.EventReceiverType,Microsoft.SharePoint.Client.EventReceiverSynchronization,System.Boolean)
Registers a remote event receiver
> ##### Parameters
> **list:** The list to process

> **name:** The name of the event receiver (needs to be unique among the event receivers registered on this list)

> **url:** The URL of the remote WCF service that handles the event

> **eventReceiverType:** 

> **synchronization:** 

> **force:** If True any event already registered with the same name will be removed first.

> ##### Return value
> Returns an EventReceiverDefinition if succeeded. Returns null if failed.

#### AddRemoteEventReceiver(Microsoft.SharePoint.Client.List,System.String,System.String,Microsoft.SharePoint.Client.EventReceiverType,Microsoft.SharePoint.Client.EventReceiverSynchronization,System.Int32,System.Boolean)
Registers a remote event receiver
> ##### Parameters
> **list:** The list to process

> **name:** The name of the event receiver (needs to be unique among the event receivers registered on this list)

> **url:** The URL of the remote WCF service that handles the event

> **eventReceiverType:** 

> **synchronization:** 

> **sequenceNumber:** 

> **force:** If True any event already registered with the same name will be removed first.

> ##### Return value
> Returns an EventReceiverDefinition if succeeded. Returns null if failed.

#### GetEventReceiverById(Microsoft.SharePoint.Client.List,System.Guid)
Returns an event receiver definition
> ##### Parameters
> **list:** 

> **id:** 

> ##### Return value
> 

#### GetEventReceiverByName(Microsoft.SharePoint.Client.List,System.String)
Returns an event receiver definition
> ##### Parameters
> **list:** The list to process

> **name:** 

> ##### Return value
> 

#### SetPropertyBagValue(Microsoft.SharePoint.Client.List,System.String,System.Int32)
Sets a key/value pair in the web property bag
> ##### Parameters
> **list:** The list to process

> **key:** Key for the property bag entry

> **value:** Integer value for the property bag entry


#### SetPropertyBagValue(Microsoft.SharePoint.Client.List,System.String,System.String)
Sets a key/value pair in the list property bag
> ##### Parameters
> **list:** List that will hold the property bag entry

> **key:** Key for the property bag entry

> **value:** String value for the property bag entry


#### SetPropertyBagValueInternal(Microsoft.SharePoint.Client.List,System.String,System.Object)
Sets a key/value pair in the list property bag
> ##### Parameters
> **list:** List that will hold the property bag entry

> **key:** Key for the property bag entry

> **value:** Value for the property bag entry


#### GetPropertyBagValueInt(Microsoft.SharePoint.Client.List,System.String,System.Int32)
Get int typed property bag value. If does not contain, returns default value.
> ##### Parameters
> **list:** List to read the property bag value from

> **key:** Key of the property bag entry to return

> **defaultValue:** 

> ##### Return value
> Value of the property bag entry as integer

#### GetPropertyBagValueString(Microsoft.SharePoint.Client.List,System.String,System.String)
Get string typed property bag value. If does not contain, returns given default value.
> ##### Parameters
> **list:** List to read the property bag value from

> **key:** Key of the property bag entry to return

> **defaultValue:** 

> ##### Return value
> Value of the property bag entry as string

#### GetPropertyBagValueInternal(Microsoft.SharePoint.Client.List,System.String)
Type independent implementation of the property gettter.
> ##### Parameters
> **list:** List to read the property bag value from

> **key:** Key of the property bag entry to return

> ##### Return value
> Value of the property bag entry

#### PropertyBagContainsKey(Microsoft.SharePoint.Client.List,System.String)
Checks if the given property bag entry exists
> ##### Parameters
> **list:** List to be processed

> **key:** Key of the property bag entry to check

> ##### Return value
> True if the entry exists, false otherwise

#### RemoveContentTypeByName(Microsoft.SharePoint.Client.List,System.String)
Removes a content type from a list/library by name
> ##### Parameters
> **list:** The list

> **contentTypeName:** The content type name to remove from the list

> ##### Exceptions
> **System.ArgumentException:** Thrown when contentTypeName is a zero-length string or contains only white space

> **System.ArgumentNullException:** contentTypeName is null


#### CreateDocumentLibrary(Microsoft.SharePoint.Client.Web,System.String,System.Boolean,System.String)
Adds a document library to a web. Execute Query is called during this implementation
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listName:** Name of the library

> **enableVersioning:** Enable versioning on the list

> **urlPath:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when listName is a zero-length string or contains only white space

> **System.ArgumentNullException:** listName is null


#### ListExists(Microsoft.SharePoint.Client.Web,System.String)
Checks if list exists on the particular site based on the list Title property.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list to be checked.

> ##### Return value
> True if the list exists
> ##### Exceptions
> **System.ArgumentException:** Thrown when listTitle is a zero-length string or contains only white space

> **System.ArgumentNullException:** listTitle is null


#### ListExists(Microsoft.SharePoint.Client.Web,System.Uri)
Checks if list exists on the particular site based on the list's site relative path.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **siteRelativeUrlPath:** Site relative path of the list

> ##### Return value
> True if the list exists

#### ListExists(Microsoft.SharePoint.Client.Web,System.Guid)
Checks if list exists on the particular site based on the list id property.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **id:** The id of the list to be checked.

> ##### Return value
> True if the list exists
> ##### Exceptions
> **System.ArgumentException:** Thrown when listTitle is a zero-length string or contains only white space

> **System.ArgumentNullException:** listTitle is null


#### CreateList(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.ListTemplateType,System.String,System.Boolean,System.Boolean,System.String,System.Boolean)
Adds a default list to a site
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listType:** Built in list template type

> **listName:** Name of the list

> **enableVersioning:** Enable versioning on the list

> **updateAndExecuteQuery:** (Optional) Perform list update and executequery, defaults to true

> **urlPath:** (Optional) URL to use for the list

> **enableContentTypes:** (Optional) Enable content type management

> ##### Return value
> The newly created list

#### CreateList(Microsoft.SharePoint.Client.Web,System.Guid,System.Int32,System.String,System.Boolean,System.Boolean,System.String,System.Boolean)
Adds a custom list to a site
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **featureId:** Feature that contains the list template

> **listType:** Type ID of the list, within the feature

> **listName:** Name of the list

> **enableVersioning:** Enable versioning on the list

> **updateAndExecuteQuery:** (Optional) Perform list update and executequery, defaults to true

> **urlPath:** (Optional) URL to use for the list

> **enableContentTypes:** (Optional) Enable content type management

> ##### Return value
> The newly created list

#### UpdateListVersioning(Microsoft.SharePoint.Client.Web,System.String,System.Boolean,System.Boolean,System.Boolean)
Enable/disable versioning on a list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listName:** List to operate on

> **enableVersioning:** True to enable versioning, false to disable

> **enableMinorVersioning:** Enable/Disable minor versioning

> **updateAndExecuteQuery:** Perform list update and executequery, defaults to true

> ##### Exceptions
> **System.ArgumentException:** Thrown when listName is a zero-length string or contains only white space

> **System.ArgumentNullException:** listName is null


#### UpdateListVersioning(Microsoft.SharePoint.Client.List,System.Boolean,System.Boolean,System.Boolean)
Enable/disable versioning on a list
> ##### Parameters
> **list:** List to be processed

> **enableVersioning:** True to enable versioning, false to disable

> **enableMinorVersioning:** Enable/Disable minor versioning

> **updateAndExecuteQuery:** Perform list update and executequery, defaults to true


#### UpdateTaxonomyFieldDefaultValue(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.Guid,System.Guid)
Sets the default value for a managed metadata column in the specified list. This operation will not change existing items in the list
> ##### Parameters
> **web:** Extension web

> **termName:** Name of a specific term

> **listName:** Name of list

> **fieldInternalName:** Internal name of field

> **groupGuid:** TermGroup Guid

> **termSetGuid:** TermSet Guid


#### SetJSLinkCustomizations(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.PageType,System.String)
Sets JS link customization for a list form
> ##### Parameters
> **list:** SharePoint list

> **pageType:** Type of form

> **jslink:** JSLink to set to the form. Set to empty string to remove the set JSLink customization. Specify multiple values separated by pipe symbol. For e.g.: ~sitecollection/_catalogs/masterpage/jquery-2.1.0.min.js|~sitecollection/_catalogs/masterpage/custom.js


#### SetJSLinkCustomizations(Microsoft.SharePoint.Client.List,System.String,System.String)
Sets JS link customization for a list view page
> ##### Parameters
> **list:** SharePoint list

> **serverRelativeUrl:** url of the view page

> **jslink:** JSLink to set to the form. Set to empty string to remove the set JSLink customization. Specify multiple values separated by pipe symbol. For e.g.: ~sitecollection/_catalogs/masterpage/jquery-2.1.0.min.js|~sitecollection/_catalogs/masterpage/custom.js


#### SetLocalizationLabelsForList(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String)
Can be used to set translations for different cultures.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list

> **cultureName:** Culture name like en-us or fi-fi

> **titleResource:** Localized Title string

> **descriptionResource:** Localized Description string

> ##### Exceptions
> **System.ArgumentException:** Thrown when listTitle, cultureName, titleResource, descriptionResource is a zero-length string or contains only white space

> **System.ArgumentNullException:** listTitle, cultureName, titleResource, descriptionResource is null


#### SetLocalizationLabelsForList(Microsoft.SharePoint.Client.List,System.String,System.String,System.String)
Can be used to set translations for different cultures.
> ##### Parameters
> **list:** List to be processed

> **cultureName:** Culture name like en-us or fi-fi

> **titleResource:** Localized Title string

> **descriptionResource:** Localized Description string

> ##### Example
> 
                list.SetLocalizationForSiteLabels("fi-fi", "Name of the site in Finnish", "Description in Finnish");
            

#### GetListID(Microsoft.SharePoint.Client.Web,System.String)
Returns the GUID id of a list
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listName:** List to operate on

> ##### Exceptions
> **System.ArgumentException:** Thrown when listName is a zero-length string or contains only white space

> **System.ArgumentNullException:** listName is null


#### GetListByTitle(Microsoft.SharePoint.Client.Web,System.String,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.List,System.Object}}[])
Get list by using Title
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **listTitle:** Title of the list to return

> **expressions:** Additional list of lambda expressions of properties to load alike l => l.BaseType

> ##### Return value
> Loaded list instance matching to title or null
> ##### Exceptions
> **System.ArgumentException:** Thrown when listTitle is a zero-length string or contains only white space

> **System.ArgumentNullException:** listTitle is null


#### GetListByUrl(Microsoft.SharePoint.Client.Web,System.String,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.List,System.Object}}[])
Get list by using Url
> ##### Parameters
> **web:** Web (site) to be processed

> **webRelativeUrl:** Url of list relative to the web (site), e.g. lists/testlist

> **expressions:** Additional list of lambda expressions of properties to load alike l => l.BaseType

> ##### Return value
> 

#### GetPagesLibrary(Microsoft.SharePoint.Client.Web)
Gets the publishing pages library of the web based on site language
> ##### Parameters
> **web:** The web.

> ##### Return value
> The publishing pages library. Returns null if library was not found.
> ##### Exceptions
> **System.InvalidOperationException:** Could not load pages library URL name from 'cmscore' resources file.


#### GetWebRelativeUrl(Microsoft.SharePoint.Client.List)
Gets the web relative URL. Allow users to get the web relative URL of a list. This is useful when exporting lists as it can then be used as a parameter to Web.GetListByUrl().
> ##### Parameters
> **list:** The list to export the URL of.

> ##### Return value
> The web relative URL of the list.

#### GetWebRelativeUrl(System.String,System.String)
Gets the web relative URL.
> ##### Parameters
> **listRootFolderServerRelativeUrl:** The list root folder server relative URL.

> **parentWebServerRelativeUrl:** The parent web server relative URL.

> ##### Return value
> The web relative URL.
> ##### Exceptions
> **System.Exception:** Cannot establish web relative URL from the list root folder URI and the parent web URI.


#### SetListPermission(Microsoft.SharePoint.Client.List,OfficeDevPnP.Core.Enums.BuiltInIdentity,Microsoft.SharePoint.Client.RoleType)
Set custom permission to the list
> ##### Parameters
> **list:** List on which permission to be set

> **user:** Built in user

> **roleType:** Role type


#### SetListPermission(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.Principal,Microsoft.SharePoint.Client.RoleType)
Set custom permission to the list
> ##### Parameters
> **list:** List on which permission to be set

> **principal:** SharePoint Group or User

> **roleType:** Role type


#### CreateViewsFromXMLFile(Microsoft.SharePoint.Client.Web,System.String,System.String)
Creates list views based on specific xml structure from file
> ##### Parameters
> **web:** 

> **listUrl:** 

> **filePath:** 


#### CreateViewsFromXMLString(Microsoft.SharePoint.Client.Web,System.String,System.String)
Creates views based on specific xml structure from string
> ##### Parameters
> **web:** 

> **listUrl:** 

> **xmlString:** 


#### CreateViewsFromXML(Microsoft.SharePoint.Client.Web,System.String,System.Xml.XmlDocument)
Create list views based on xml structure loaded to memory
> ##### Parameters
> **web:** 

> **listUrl:** 

> **xmlDoc:** 


#### CreateViewsFromXMLFile(Microsoft.SharePoint.Client.List,System.String)
Create list views based on specific xml structure in external file
> ##### Parameters
> **list:** 

> **filePath:** 


#### CreateViewsFromXMLString(Microsoft.SharePoint.Client.List,System.String)
Create list views based on specific xml structure in string
> ##### Parameters
> **list:** 

> **xmlString:** 


#### CreateViewsFromXML(Microsoft.SharePoint.Client.List,System.Xml.XmlDocument)
Actual implementation of the view creation logic based on given xml
> ##### Parameters
> **list:** 

> **xmlDoc:** 


#### CreateView(Microsoft.SharePoint.Client.List,System.String,Microsoft.SharePoint.Client.ViewType,System.String[],System.UInt32,System.Boolean,System.String,System.Boolean,System.Boolean)
Create view to existing list
> ##### Parameters
> **list:** 

> **viewName:** 

> **viewType:** 

> **viewFields:** 

> **rowLimit:** 

> **setAsDefault:** 

> **query:** 

> **personal:** 

> **paged:** 


#### GetViewById(Microsoft.SharePoint.Client.List,System.Guid,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.View,System.Object}}[])
Gets a view by Id
> ##### Parameters
> **list:** 

> **id:** 

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> returns null if not found

#### GetViewByName(Microsoft.SharePoint.Client.List,System.String,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.View,System.Object}}[])
Gets a view by Name
> ##### Parameters
> **list:** 

> **name:** 

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> returns null if not found

#### SetDefaultColumnValues(Microsoft.SharePoint.Client.List,System.Collections.Generic.IEnumerable{OfficeDevPnP.Core.Entities.IDefaultColumnValue})
Sets default values for column values. In order to for instance set the default Enterprise Metadata keyword field to a term, add the enterprise metadata keyword to a library (internal name "TaxKeyword") Column values are defined by the DefaultColumnValue class that has 3 properties: RelativeFolderPath : / to set a default value for the root of the document library, or /foldername to specify a subfolder FieldInternalName : The name of the field to set. For instance "TaxKeyword" to set the Enterprise Metadata field Terms : A collection of Taxonomy terms to set Supported column types: Metadata, Text, Choice, MultiChoice, People, Boolean, DateTime, Number, Currency
> ##### Parameters
> **list:** 

> **columnValues:** 


#### GetDefaultColumnValues(Microsoft.SharePoint.Client.List)
Gets default values for column values. The returned list contains one dictionary per default setting per folder. Each dictionary has the following keys set: Path, Field, Value Path: Relative path to the library/folder Field: Internal name of the field which has a default value Value: The default value for the field
> ##### Parameters
> **list:** 


#### ReIndexList(Microsoft.SharePoint.Client.List)
Queues a list for a full crawl the next incremental crawl
> ##### Parameters
> **list:** 


## SharePoint.Client.NavigationExtensions
            
This class holds deprecated navigation related methods
            
This class holds navigation related methods
        
### Methods


#### GetNavigationSettings(Microsoft.SharePoint.Client.Web)
Returns the navigation settings for the selected web
> ##### Parameters
> **web:** 

> ##### Return value
> 

#### UpdateNavigationSettings(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.AreaNavigationEntity)
Updates navigation settings for the current web
> ##### Parameters
> **web:** 

> **navigationSettings:** 


#### GetEditableNavigationTermSet(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.ManagedNavigationKind)
Returns an editable version of the Global Navigation TermSet for a web site
> ##### Parameters
> **web:** The target web.

> **navigationKind:** Declares whether to look for Current or Global Navigation

> ##### Return value
> The editable Global Navigation TermSet

#### IsManagedNavigationEnabled(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.ManagedNavigationKind)
Determines whether the current Web has the managed navigation enabled
> ##### Parameters
> **web:** The target web.

> **navigationKind:** The kind of navigation (Current or Global).

> ##### Return value
> A boolean result of the test.

#### AddNavigationNode(Microsoft.SharePoint.Client.Web,System.String,System.Uri,System.String,OfficeDevPnP.Core.Enums.NavigationType,System.Boolean,System.Boolean)
Add a node to quick launch, top navigation bar or search navigation. The node will be added as the last node in the collection.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **nodeTitle:** the title of node to add

> **nodeUri:** the url of node to add

> **parentNodeTitle:** if string.Empty, then will add this node as top level node

> **navigationType:** the type of navigation, quick launch, top navigation or search navigation

> **isExternal:** true if the link is an external link

> **asLastNode:** true if the link should be added as the last node of the collection

> ##### Return value
> Newly added NavigationNode

#### DeleteNavigationNode(Microsoft.SharePoint.Client.Web,System.String,System.String,OfficeDevPnP.Core.Enums.NavigationType)
Deletes a navigation node from the quickLaunch or top navigation bar
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **nodeTitle:** the title of node to delete

> **parentNodeTitle:** if string.Empty, then will delete this node as top level node

> **navigationType:** the type of navigation, quick launch, top navigation or search navigation


#### DeleteAllNavigationNodes(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Enums.NavigationType)
Deletes all Navigation Nodes from a given navigation
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **navigationType:** The type of navigation to support


#### UpdateNavigationInheritance(Microsoft.SharePoint.Client.Web,System.Boolean)
Updates the navigation inheritance setting
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **inheritNavigation:** boolean indicating if navigation inheritance is needed or not


#### LoadSearchNavigation(Microsoft.SharePoint.Client.Web)
Loads the search navigation nodes
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> ##### Return value
> Collection of NavigationNode instances

#### AddCustomAction(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.CustomActionEntity)
Adds custom action to a web. If the CustomAction exists the item will be updated. Setting CustomActionEntity.Remove == true will delete the CustomAction.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **customAction:** Information about the custom action be added or deleted

> ##### Return value
> True if action was successfull
> ##### Example
> 
            var editAction = new CustomActionEntity()
            {
               Title = "Edit Site Classification",
               Description = "Manage business impact information for site collection or sub sites.",
               Sequence = 1000,
               Group = "SiteActions",
               Location = "Microsoft.SharePoint.StandardMenu",
               Url = EditFormUrl,
               ImageUrl = EditFormImageUrl,
               Rights = new BasePermissions(),
            };
            editAction.Rights.Set(PermissionKind.ManageWeb);
            web.AddCustomAction(editAction);
            

#### AddCustomAction(Microsoft.SharePoint.Client.Site,OfficeDevPnP.Core.Entities.CustomActionEntity)
Adds custom action to a site collection. If the CustomAction exists the item will be updated. Setting CustomActionEntity.Remove == true will delete the CustomAction.
> ##### Parameters
> **site:** Site collection to be processed

> **customAction:** Information about the custom action be added or deleted

> ##### Return value
> True if action was successfull

#### GetCustomActions(Microsoft.SharePoint.Client.Web,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.UserCustomAction,System.Object}}[])
Returns all custom actions in a web
> ##### Parameters
> **web:** The web to process

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> 

#### GetCustomActions(Microsoft.SharePoint.Client.Site,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.UserCustomAction,System.Object}}[])
Returns all custom actions in a web
> ##### Parameters
> **site:** The site to process

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> 

#### DeleteCustomAction(Microsoft.SharePoint.Client.Web,System.Guid)
Removes a custom action
> ##### Parameters
> **web:** The web to process

> **id:** The id of the action to remove. GetCustomActions


#### DeleteCustomAction(Microsoft.SharePoint.Client.Site,System.Guid)
Removes a custom action
> ##### Parameters
> **site:** The site to process

> **id:** The id of the action to remove. GetCustomActions


#### CustomActionExists(Microsoft.SharePoint.Client.Web,System.String)
Utility method to check particular custom action already exists on the web
> ##### Parameters
> **web:** 

> **name:** Name of the custom action

> ##### Return value
> 

#### CustomActionExists(Microsoft.SharePoint.Client.Site,System.String)
Utility method to check particular custom action already exists on the web
> ##### Parameters
> **site:** 

> **name:** Name of the custom action

> ##### Return value
> 

## SharePoint.Client.PageExtensions
            
Class that holds all deprecated page and web part related operations
            
Class that handles all page and web part related operations
        
### Methods


#### GetWikiPageContent(Microsoft.SharePoint.Client.Web,System.String)
Returns the HTML contents of a wiki page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **serverRelativePageUrl:** Server relative url of the page, e.g. /sites/demo/SitePages/Test.aspx

> ##### Return value
> 
> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl is null


#### GetWebParts(Microsoft.SharePoint.Client.Web,System.String)
List the web parts on a page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **serverRelativePageUrl:** Server relative url of the page containing the webparts

> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl is null


#### AddWebPartToWebPartPage(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.WebPartEntity,System.String)
Inserts a web part on a web part page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **webPart:** Information about the web part to insert

> **page:** Page to add the web part on

> ##### Return value
> Returns the added object
> ##### Exceptions
> **System.ArgumentException:** Thrown when page is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when webPart or page is null


#### AddWebPartToWebPartPage(Microsoft.SharePoint.Client.Web,System.String,OfficeDevPnP.Core.Entities.WebPartEntity)
Inserts a web part on a web part page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **serverRelativePageUrl:** Page to add the web part on

> **webPart:** Information about the web part to insert

> ##### Return value
> Returns the added object
> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl or webPart is null


#### AddWebPartToWikiPage(Microsoft.SharePoint.Client.Web,System.String,OfficeDevPnP.Core.Entities.WebPartEntity,System.String,System.Int32,System.Int32,System.Boolean)
Add web part to a wiki style page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **folder:** System name of the wiki page library - typically sitepages

> **webPart:** Information about the web part to insert

> **page:** Page to add the web part on

> **row:** Row of the wiki table that should hold the inserted web part

> **col:** Column of the wiki table that should hold the inserted web part

> **addSpace:** Does a blank line need to be added after the web part (to space web parts)

> ##### Return value
> Returns the added object
> ##### Exceptions
> **System.ArgumentException:** Thrown when folder or page is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when folder, webPart or page is null


#### AddWebPartToWikiPage(Microsoft.SharePoint.Client.Web,System.String,OfficeDevPnP.Core.Entities.WebPartEntity,System.Int32,System.Int32,System.Boolean)
Add web part to a wiki style page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **serverRelativePageUrl:** Server relative url of the page to add the webpart to

> **webPart:** Information about the web part to insert

> **row:** Row of the wiki table that should hold the inserted web part

> **col:** Column of the wiki table that should hold the inserted web part

> **addSpace:** Does a blank line need to be added after the web part (to space web parts)

> ##### Return value
> Returns the added object
> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl or webPart is null


#### AddLayoutToWikiPage(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.WikiPageLayout,System.String)
Applies a layout to a wiki page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **layout:** Wiki page layout to be applied

> **serverRelativePageUrl:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl is null


#### AddLayoutToWikiPage(Microsoft.SharePoint.Client.Web,System.String,OfficeDevPnP.Core.WikiPageLayout,System.String)
Applies a layout to a wiki page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **folder:** System name of the wiki page library - typically sitepages

> **layout:** Wiki page layout to be applied

> **page:** Name of the page that will get a new wiki page layout

> ##### Exceptions
> **System.ArgumentException:** Thrown when folder or page is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when folder or page is null


#### AddHtmlToWikiPage(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String)
Add html to a wiki style page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **folder:** System name of the wiki page library - typically sitepages

> **html:** The html to insert

> **page:** Page to add the html on

> ##### Exceptions
> **System.ArgumentException:** Thrown when folder, html or page is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when folder, html or page is null


#### AddHtmlToWikiPage(Microsoft.SharePoint.Client.Web,System.String,System.String)
Add HTML to a wiki page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **serverRelativePageUrl:** 

> **html:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl or html is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl or html is null


#### AddHtmlToWikiPage(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.Int32,System.Int32)
Add a HTML fragment to a location on a wiki style page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **folder:** System name of the wiki page library - typically sitepages

> **html:** html to be inserted

> **page:** Page to add the web part on

> **row:** Row of the wiki table that should hold the inserted web part

> **col:** Column of the wiki table that should hold the inserted web part

> ##### Exceptions
> **System.ArgumentException:** Thrown when folder, html or page is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when folder, html or page is null


#### AddHtmlToWikiPage(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Int32,System.Int32)
Add a HTML fragment to a location on a wiki style page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **serverRelativePageUrl:** server relative Url of the page to add the fragment to

> **html:** html to be inserted

> **row:** Row of the wiki table that should hold the inserted web part

> **col:** Column of the wiki table that should hold the inserted web part

> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl or html is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl or html is null


#### DeleteWebPart(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String)
Deletes a web part from a page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **folder:** System name of the wiki page library - typically sitepages

> **title:** Title of the web part that needs to be deleted

> **page:** Page to remove the web part from

> ##### Exceptions
> **System.ArgumentException:** Thrown when folder, title or page is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when folder, title or page is null


#### DeleteWebPart(Microsoft.SharePoint.Client.Web,System.String,System.String)
Deletes a web part from a page
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **serverRelativePageUrl:** Server relative URL of the page to remove

> **title:** Title of the web part that needs to be deleted

> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl or title is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl or title is null


#### AddClientSidePage(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Adds a client side "modern" page to a "classic" or "modern" site
> ##### Parameters
> **web:** Web to add the page to

> **pageName:** Name (e.g. demo.aspx) of the page to be added

> **alreadyPersist:** Already persist the created, empty, page before returning the instantiated instance

> ##### Return value
> A instance

#### LoadClientSidePage(Microsoft.SharePoint.Client.Web,System.String)
Loads a client side "modern" page
> ##### Parameters
> **web:** Web to load the page from

> **pageName:** Name (e.g. demo.aspx) of the page to be loaded

> ##### Return value
> A instance

#### AddWikiPage(Microsoft.SharePoint.Client.Web,System.String,System.String)
Adds a blank Wiki page to the site pages library
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **wikiPageLibraryName:** Name of the wiki page library

> **wikiPageName:** Wiki page to operate on

> ##### Return value
> The relative URL of the added wiki page
> ##### Exceptions
> **System.ArgumentException:** Thrown when wikiPageLibraryName or wikiPageName is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when wikiPageLibraryName or wikiPageName is null


#### EnsureWikiPage(Microsoft.SharePoint.Client.Web,System.String,System.String)
Returns the Url for the requested wiki page, creates it if the pageis not yet available
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **wikiPageLibraryName:** Name of the wiki page library

> **wikiPageName:** Wiki page to operate on

> ##### Return value
> The relative URL of the added wiki page
> ##### Exceptions
> **System.ArgumentException:** Thrown when wikiPageLibraryName or wikiPageName is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when wikiPageLibraryName or wikiPageName is null


#### AddWikiPageByUrl(Microsoft.SharePoint.Client.Web,System.String,System.String)
Adds a wiki page by Url
> ##### Parameters
> **web:** The web to process

> **serverRelativePageUrl:** Server relative URL of the wiki page to process

> **html:** HTML to add to wiki page

> ##### Exceptions
> **System.ArgumentException:** Thrown when serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when serverRelativePageUrl is null


#### SetWebPartProperty(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Guid,System.String)
Sets a web part property
> ##### Parameters
> **web:** The web to process

> **key:** The key to update

> **value:** The value to set

> **id:** The id of the webpart

> **serverRelativePageUrl:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when key or serverRelativePageUrl is null


#### SetWebPartProperty(Microsoft.SharePoint.Client.Web,System.String,System.Int32,System.Guid,System.String)
Sets a web part property
> ##### Parameters
> **web:** The web to process

> **key:** The key to update

> **value:** The value to set

> **id:** The id of the webpart

> **serverRelativePageUrl:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when key or serverRelativePageUrl is null


#### SetWebPartProperty(Microsoft.SharePoint.Client.Web,System.String,System.Boolean,System.Guid,System.String)
Sets a web part property
> ##### Parameters
> **web:** The web to process

> **key:** The key to update

> **value:** The value to set

> **id:** The id of the webpart

> **serverRelativePageUrl:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when key or serverRelativePageUrl is null


#### GetWebPartProperties(Microsoft.SharePoint.Client.Web,System.Guid,System.String)
Returns web part properties
> ##### Parameters
> **web:** The web to process

> **id:** The id of the webpart

> **serverRelativePageUrl:** 

> ##### Exceptions
> **System.ArgumentException:** Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space

> **System.ArgumentNullException:** Thrown when key or serverRelativePageUrl is null


#### AddNavigationFriendlyUrl(Microsoft.SharePoint.Client.Publishing.PublishingPage,Microsoft.SharePoint.Client.Web,System.String,System.String,Microsoft.SharePoint.Client.Publishing.Navigation.NavigationTermSetItem,System.Boolean,System.Boolean)
Adds a user-friendly URL for a PublishingPage object.
> ##### Parameters
> **page:** The target page to add to managed navigation.

> **web:** The target web.

> **navigationTitle:** The title for the navigation item.

> **friendlyUrlSegment:** The user-friendly text to use as the URL segment.

> **editableParent:** The parent NavigationTermSetItem object below which this new friendly URL should be created.

> **showInGlobalNavigation:** Defines whether the navigation item has to be shown in the Global Navigation, optional and default to true.

> **showInCurrentNavigation:** Defines whether the navigation item has to be shown in the Current Navigation, optional and default to true.

> ##### Return value
> The simple link URL just created.

## SharePoint.Client.ProvisioningExtensions
            
File-based (CAML) deprecated provisioning extensions
            
File-based (CAML) provisioning extensions
        
### Methods


#### ProvisionElementFile(Microsoft.SharePoint.Client.Web,System.String)
Provisions the items defined by the specified Elements (CAML) file; currently only supports modules (files).
> ##### Parameters
> **web:** Web to provision the elements to

> **path:** Path to the XML file containing the Elements CAML defintion


#### ProvisionElementXml(Microsoft.SharePoint.Client.Web,System.String,System.Xml.Linq.XElement)
Provisions the items defined by the specified Elements (CAML) XML; currently only supports modules (files).
> ##### Parameters
> **web:** Web to provision the elements to

> **baseFolder:** Base local folder to find any referenced items, e.g. files

> **elementsXml:** Elements (CAML) XML element that defines the items to provision; currently only supports modules (files)


#### ProvisionModuleInternal(Microsoft.SharePoint.Client.Web,System.String,System.Xml.Linq.XElement)
Uploads all files defined by the moduleXml

#### ProvisionFileInternal(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Xml.Linq.XElement,System.Boolean)
Uploads the file defined by the fileXml, creating folders as necessary.

## SharePoint.Client.RecordsManagementExtensions
            
Class that deals with deprecated records management functionality
            
Class that deals with records management functionality
        
### Fields

#### INPLACE_RECORDS_MANAGEMENT_FEATURE_ID
Defines the ID of the Inplace Records Management Feature
#### ECM_SITE_RECORD_DECLARATION_DEFAULT
Defines the name of the ECM Site Record Declaration Default propertybag value
#### ECM_SITE_RECORD_RESTRICTIONS
Defines the name of the ECM Site Record Restrictions propertybag value
#### ECM_SITE_RECORD_DECLARATION_BY
Defines the name of the ECM Site Record Declaration by propertybag value
#### ECM_SITE_RECORD_UNDECLARATION_BY
Defines the name of the ECM Site Record Undeclaration by propertybag value
#### ECM_ALLOW_MANUAL_DECLARATION
Defines the name of the ECM Allow Manual Declaration propertybag value
#### ECM_IPR_LIST_USE_LIST_SPECIFIC
Defines the name of the ECM IPR List use List Specific propertybag value
#### ECM_AUTO_DECLARE_RECORDS
Defines the name of the ECM auto declare records propertybag value
### Methods


#### IsInPlaceRecordsManagementActive(Microsoft.SharePoint.Client.Site)
Checks if in place records management functionality is enabled for this site collection
> ##### Parameters
> **site:** Site collection to operate on

> ##### Return value
> True if in place records management is enabled, false otherwise

#### ActivateInPlaceRecordsManagementFeature(Microsoft.SharePoint.Client.Site)
Activate the in place records management feature
> ##### Parameters
> **site:** Site collection to operate on


#### DisableInPlaceRecordsManagementFeature(Microsoft.SharePoint.Client.Site)
Deactivate the in place records management feature
> ##### Parameters
> **site:** Site collection to operate on


#### EnableSiteForInPlaceRecordsManagement(Microsoft.SharePoint.Client.Site)
Enable in place records management. The in place records management feature will be enabled and the in place record management will be enabled in all locations with record declaration allowed by all contributors and undeclaration by site admins
> ##### Parameters
> **site:** Site collection to operate on


#### SetManualRecordDeclarationInAllLocations(Microsoft.SharePoint.Client.Site,System.Boolean)
Defines if in place records management is allowed in all places
> ##### Parameters
> **site:** Site collection to operate on

> **inAllPlaces:** True if allowed in all places, false otherwise


#### GetManualRecordDeclarationInAllLocations(Microsoft.SharePoint.Client.Site)
Get the value of the records management is allowed in all places setting
> ##### Parameters
> **site:** Site collection to operate on

> ##### Return value
> True if records management is allowed in all places, false otherwise

#### SetRecordRestrictions(Microsoft.SharePoint.Client.Site,OfficeDevPnP.Core.EcmSiteRecordRestrictions)
Defines the restrictions that are placed on a document once it's declared as a record
> ##### Parameters
> **site:** Site collection to operate on

> **restrictions:** enum that holds the restrictions to be applied


#### GetRecordRestrictions(Microsoft.SharePoint.Client.Site)
Gets the current restrictions on declared records
> ##### Parameters
> **site:** Site collection to operate on

> ##### Return value
> enum that defines the current restrictions

#### SetRecordDeclarationBy(Microsoft.SharePoint.Client.Site,OfficeDevPnP.Core.EcmRecordDeclarationBy)
Defines who can declare records
> ##### Parameters
> **site:** Site collection to operate on

> **by:** enum that defines who can declare a record


#### GetRecordDeclarationBy(Microsoft.SharePoint.Client.Site)
Gets who can declare records
> ##### Parameters
> **site:** Site collection to operate on

> ##### Return value
> enum that defines who can declare a record

#### SetRecordUnDeclarationBy(Microsoft.SharePoint.Client.Site,OfficeDevPnP.Core.EcmRecordDeclarationBy)
Defines who can undeclare records
> ##### Parameters
> **site:** Site collection to operate on

> **by:** enum that defines who can undeclare a record


#### GetRecordUnDeclarationBy(Microsoft.SharePoint.Client.Site)
Gets who can undeclare records
> ##### Parameters
> **site:** Site collection to operate on

> ##### Return value
> enum that defines who can undeclare a record

#### IsListRecordSettingDefined(Microsoft.SharePoint.Client.List)
Checks if this list has active in place records management settings defined
> ##### Parameters
> **list:** List to operate against

> ##### Return value
> True if in place records management settings are active for this list

#### SetListManualRecordDeclaration(Microsoft.SharePoint.Client.List,OfficeDevPnP.Core.EcmListManualRecordDeclaration)
Defines the manual in place record declaration for this list
> ##### Parameters
> **list:** List to operate against

> **settings:** enum that defines the manual in place record declaration settings for this list


#### GetListManualRecordDeclaration(Microsoft.SharePoint.Client.List)
Gets the manual in place record declaration for this list
> ##### Parameters
> **list:** List to operate against

> ##### Return value
> enum that defines the manual in place record declaration settings for this list

#### SetListAutoRecordDeclaration(Microsoft.SharePoint.Client.List,System.Boolean)
Defines if auto record declaration is active for this list: all added items will be automatically declared as a record if active
> ##### Parameters
> **list:** List to operate on

> **autoDeclareRecords:** True to automatically declare all added items as record, false otherwise


#### GetListAutoRecordDeclaration(Microsoft.SharePoint.Client.List)
Returns if auto record declaration is active for this list
> ##### Parameters
> **list:** List to operate against

> ##### Return value
> True if auto record declaration is active, false otherwise

## SharePoint.Client.SearchExtensions
            
Class for deprecated search extension methods
            
Class for Search extension methods
        
### Methods


#### ExportSearchSettings(Microsoft.SharePoint.Client.ClientContext,System.String,Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel)
Exports the search settings to file.
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **exportFilePath:** Path where to export the search settings

> **searchSettingsExportLevel:** Search settings export level Reference: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.administration.searchobjectlevel(v=office.15).aspx


#### GetSearchConfiguration(Microsoft.SharePoint.Client.Web)
Returns the current search configuration as as string
> ##### Parameters
> **web:** 

> ##### Return value
> 

#### GetSearchConfiguration(Microsoft.SharePoint.Client.Site)
Returns the current search configuration as as string
> ##### Parameters
> **site:** 

> ##### Return value
> 

#### GetSearchConfigurationImplementation(Microsoft.SharePoint.Client.ClientRuntimeContext,Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel)
Returns the current search configuration for the specified object level
> ##### Parameters
> **context:** 

> **searchSettingsObjectLevel:** 

> ##### Return value
> 

#### ImportSearchSettings(Microsoft.SharePoint.Client.ClientContext,System.String,Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel)
Imports search settings from file.
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **searchSchemaImportFilePath:** Search schema xml file path

> **searchSettingsImportLevel:** Search settings import level Reference: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.search.administration.searchobjectlevel(v=office.15).aspx


#### SetSearchConfiguration(Microsoft.SharePoint.Client.Web,System.String)
Sets the search configuration
> ##### Parameters
> **web:** 

> **searchConfiguration:** 


#### SetSearchConfiguration(Microsoft.SharePoint.Client.Site,System.String)
Sets the search configuration
> ##### Parameters
> **site:** 

> **searchConfiguration:** 


#### SetSearchConfigurationImplementation(Microsoft.SharePoint.Client.ClientRuntimeContext,Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel,System.String)
Sets the search configuration at the specified object level
> ##### Parameters
> **context:** 

> **searchObjectLevel:** 

> **searchConfiguration:** 


#### SetSiteCollectionSearchCenterUrl(Microsoft.SharePoint.Client.Web,System.String)
Sets the search center url on site collection (Site Settings -> Site collection administration --> Search Settings)
> ##### Parameters
> **web:** SharePoint site - root web

> **searchCenterUrl:** Search center url


#### GetSiteCollectionSearchCenterUrl(Microsoft.SharePoint.Client.Web)
Get the search center url for the site collection (Site Settings -> Site collection administration --> Search Settings)
> ##### Parameters
> **web:** SharePoint site - root web

> ##### Return value
> Search center url for web

#### SetWebSearchCenterUrl(Microsoft.SharePoint.Client.Web,System.String)
Sets the search results page url on current web (Site Settings -> Search --> Search Settings)
> ##### Parameters
> **web:** SharePoint current web

> **searchCenterUrl:** Search results page url


#### GetWebSearchCenterUrl(Microsoft.SharePoint.Client.Web)
Get the search results page url for the web (Site Settings -> Search --> Search Settings)
> ##### Parameters
> **web:** SharePoint site - current web

> ##### Return value
> Search results page url for web

## SharePoint.Client.SecurityExtensions
            
This manager class holds deprecated security related methods
            
This manager class holds security related methods
        
### Fields

#### MockupUserEmailCache
A dictionary to cache resolved user emails. key: user login name, value: user email. *** Don't use this cache in a real world application. *** Instead it should be replaced by a real cache with ref object to clear up intermediate records periodically.
#### MockupGroupCache
A dictionary to cache all user entities of a given SharePoint group. key: group login name, value: an array of user entities belongs to the group. *** Don't use this cache in a real world application. *** Instead it should be replaced by a real cache with ref object to clear up intermediate records periodically.
### Methods


#### GetAdministrators(Microsoft.SharePoint.Client.Web)
Get a list of site collection administrators
> ##### Parameters
> **web:** Site to operate on

> ##### Return value
> List of objects

#### AddAdministrators(Microsoft.SharePoint.Client.Web,System.Collections.Generic.List{OfficeDevPnP.Core.Entities.UserEntity},System.Boolean)
Add a site collection administrator to a site collection
> ##### Parameters
> **web:** Site to operate on

> **adminLogins:** Array of admins loginnames to add

> **addToOwnersGroup:** Optionally the added admins can also be added to the Site owners group


#### RemoveAdministrator(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.UserEntity)
Removes an administrators from the site collection
> ##### Parameters
> **web:** Site to operate on

> **admin:** that describes the admin to be removed


#### AddReaderAccess(Microsoft.SharePoint.Client.Web)
Add read access to the group "Everyone except external users".
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site


#### AddReaderAccess(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Enums.BuiltInIdentity)
Add read access to the group "Everyone except external users".
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **user:** Built in user to add to the visitors group


#### GetSharingCapabilitiesTenant(Microsoft.SharePoint.Client.Web,System.Uri)
Get the external sharing settings for the provided site. Only works in Office 365 Multi-Tenant
> ##### Parameters
> **web:** Tenant administration web

> **siteUrl:** Site to get the sharing capabilities from

> ##### Return value
> Sharing capabilities of the site collection

#### GetExternalUsersTenant(Microsoft.SharePoint.Client.Web)
Returns a list all external users in your tenant
> ##### Parameters
> **web:** Tenant administration web

> ##### Return value
> A list of objects

#### GetExternalUsersForSiteTenant(Microsoft.SharePoint.Client.Web,System.Uri)
Returns a list all external users for a given site that have at least the viewpages permission
> ##### Parameters
> **web:** Tenant administration web

> **siteUrl:** Url of the site fetch the external users for

> ##### Return value
> A list of objects

#### GetGroupID(Microsoft.SharePoint.Client.Web,System.String)
Returns the integer ID for a given group name
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **groupName:** SharePoint group name

> ##### Return value
> Integer group ID

#### AddGroup(Microsoft.SharePoint.Client.Web,System.String,System.String,System.Boolean,System.Boolean,System.Boolean)
Adds a group
> ##### Parameters
> **web:** Site to add the group to

> **groupName:** Name of the group

> **groupDescription:** Description of the group

> **groupIsOwner:** Sets the created group as group owner if true

> **updateAndExecuteQuery:** Set to false to postpone the executequery call

> **onlyAllowMembersViewMembership:** Set whether members are allowed to see group membership, defaults to false

> ##### Return value
> The created group

#### AssociateDefaultGroups(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Group,Microsoft.SharePoint.Client.Group,Microsoft.SharePoint.Client.Group)
Associate the provided groups as default owners, members or visitors groups. If a group is null then the association is not done
> ##### Parameters
> **web:** Site to operate on

> **owners:** Owners group

> **members:** Members group

> **visitors:** Visitors group


#### AddUserToGroup(Microsoft.SharePoint.Client.Web,System.String,System.String)
Adds a user to a group
> ##### Parameters
> **web:** web to operate against

> **groupName:** Name of the group

> **userLoginName:** Loginname of the user


#### AddUserToGroup(Microsoft.SharePoint.Client.Web,System.Int32,System.String)
Adds a user to a group
> ##### Parameters
> **web:** web to operate against

> **groupId:** Id of the group

> **userLoginName:** Login name of the user


#### AddUserToGroup(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Group,Microsoft.SharePoint.Client.User)
Adds a user to a group
> ##### Parameters
> **web:** Web to operate against

> **group:** Group object representing the group

> **user:** User object representing the user


#### AddUserToGroup(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Group,System.String)
Adds a user to a group
> ##### Parameters
> **web:** Web to operate against

> **group:** Group object representing the group

> **userLoginName:** Login name of the user


#### AddPermissionLevelToUser(Microsoft.SharePoint.Client.SecurableObject,System.String,Microsoft.SharePoint.Client.RoleType,System.Boolean)
Add a permission level (e.g.Contribute, Reader,...) to a user
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **userLoginName:** Loginname of the user

> **permissionLevel:** Permission level to add

> **removeExistingPermissionLevels:** Set to true to remove all other permission levels for that user


#### AddPermissionLevelToUser(Microsoft.SharePoint.Client.SecurableObject,System.String,System.String,System.Boolean)
Add a role definition (e.g.Contribute, Read, Approve) to a user
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **userLoginName:** Loginname of the user

> **roleDefinitionName:** Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Hierarchy|Restricted Read. Use the correct name of the language of the root site you are using

> **removeExistingPermissionLevels:** Set to true to remove all other permission levels for that user


#### AddPermissionLevelToGroup(Microsoft.SharePoint.Client.SecurableObject,System.String,Microsoft.SharePoint.Client.RoleType,System.Boolean)
Add a permission level (e.g.Contribute, Reader,...) to a group
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **groupName:** Name of the group

> **permissionLevel:** Permission level to add

> **removeExistingPermissionLevels:** Set to true to remove all other permission levels for that group


#### AddPermissionLevelToPrincipal(Microsoft.SharePoint.Client.SecurableObject,Microsoft.SharePoint.Client.Principal,Microsoft.SharePoint.Client.RoleType,System.Boolean)
Add a permission level (e.g.Contribute, Reader,...) to a group
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **principal:** Principal to add permission to

> **permissionLevel:** Permission level to add

> **removeExistingPermissionLevels:** Set to true to remove all other permission levels for that group


#### AddPermissionLevelToGroup(Microsoft.SharePoint.Client.SecurableObject,System.String,System.String,System.Boolean)
Add a role definition (e.g.Contribute, Read, Approve) to a group
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **groupName:** Name of the group

> **roleDefinitionName:** Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Hierarchy|Restricted Read. Use the correct name of the language of the root site you are using

> **removeExistingPermissionLevels:** Set to true to remove all other permission levels for that group


#### AddPermissionLevelToPrincipal(Microsoft.SharePoint.Client.SecurableObject,Microsoft.SharePoint.Client.Principal,System.String,System.Boolean)
Add a role definition (e.g.Contribute, Read, Approve) to a group
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **principal:** Principal to add permission to

> **roleDefinitionName:** Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Hierarchy|Restricted Read. Use the correct name of the language of the root site you are using

> **removeExistingPermissionLevels:** Set to true to remove all other permission levels for that group


#### RemovePermissionLevelFromUser(Microsoft.SharePoint.Client.SecurableObject,System.String,Microsoft.SharePoint.Client.RoleType,System.Boolean)
Removes a permission level from a user
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **userLoginName:** Loginname of user

> **permissionLevel:** Permission level to remove. If null all permission levels are removed

> **removeAllPermissionLevels:** Set to true to remove all permission level.


#### RemovePermissionLevelFromPrincipal(Microsoft.SharePoint.Client.SecurableObject,Microsoft.SharePoint.Client.Principal,Microsoft.SharePoint.Client.RoleType,System.Boolean)
Removes a permission level from a user
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **principal:** Principal to remove permission from

> **permissionLevel:** Permission level to remove. If null all permission levels are removed

> **removeAllPermissionLevels:** Set to true to remove all permission level.


#### RemovePermissionLevelFromUser(Microsoft.SharePoint.Client.SecurableObject,System.String,System.String,System.Boolean)
Removes a permission level from a user
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **userLoginName:** Loginname of user

> **roleDefinitionName:** Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Heirarchy|Restricted Read. Use the correct name of the language of the site you are using

> **removeAllPermissionLevels:** Set to true to remove all permission level.


#### RemovePermissionLevelFromPrincipal(Microsoft.SharePoint.Client.SecurableObject,Microsoft.SharePoint.Client.Principal,System.String,System.Boolean)
Removes a permission level from a user
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **principal:** Principal to remove permission from

> **roleDefinitionName:** Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Heirarchy|Restricted Read. Use the correct name of the language of the site you are using

> **removeAllPermissionLevels:** Set to true to remove all permission level.


#### RemovePermissionLevelFromGroup(Microsoft.SharePoint.Client.SecurableObject,System.String,Microsoft.SharePoint.Client.RoleType,System.Boolean)
Removes a permission level from a group
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **groupName:** name of the group

> **permissionLevel:** Permission level to remove. If null all permission levels are removed

> **removeAllPermissionLevels:** Set to true to remove all permission level.


#### RemovePermissionLevelFromGroup(Microsoft.SharePoint.Client.SecurableObject,System.String,System.String,System.Boolean)
Removes a permission level from a group
> ##### Parameters
> **securableObject:** Web/List/Item to operate against

> **groupName:** name of the group

> **roleDefinitionName:** Name of the role definition to add, Full Control|Design|Contribute|Read|Approve|Manage Heirarchy|Restricted Read. Use the correct name of the language of the site you are using

> **removeAllPermissionLevels:** Set to true to remove all permission level.


#### RemoveUserFromGroup(Microsoft.SharePoint.Client.Web,System.String,System.String)
Removes a user from a group
> ##### Parameters
> **web:** Web to operate against

> **groupName:** Name of the group

> **userLoginName:** Loginname of the user


#### RemoveUserFromGroup(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Group,Microsoft.SharePoint.Client.User)
Removes a user from a group
> ##### Parameters
> **web:** Web to operate against

> **group:** Group object to operate against

> **user:** User object that needs to be removed


#### RemoveGroup(Microsoft.SharePoint.Client.Web,System.String)
Remove a group
> ##### Parameters
> **web:** Web to operate against

> **groupName:** Name of the group


#### RemoveGroup(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Group)
Remove a group
> ##### Parameters
> **web:** Web to operate against

> **group:** Group object to remove


#### IsUserInGroup(Microsoft.SharePoint.Client.Web,System.String,System.String)
Checks if a user is member of a group
> ##### Parameters
> **web:** Web to operate against

> **groupName:** Name of the group

> **userLoginName:** Loginname of the user

> ##### Return value
> True if the user is in the group, false otherwise

#### GroupExists(Microsoft.SharePoint.Client.Web,System.String)
Checks if a group exists
> ##### Parameters
> **web:** Web to operate against

> **groupName:** Name of the group

> ##### Return value
> True if the group exists, false otherwise

#### GetAuthenticationRealm(Microsoft.SharePoint.Client.Web)
Returns the authentication realm for the current web
> ##### Parameters
> **web:** 

> ##### Return value
> 

#### GetPath(Microsoft.SharePoint.Client.SecurableObject)
Get URL path of a securable object
> ##### Parameters
> **obj:** A securable object which could be a web, a list, a list item, a document library or a document

> ##### Return value
> The URL of the securable object

#### Preload(Microsoft.SharePoint.Client.SecurableObject,System.Int32)
Load properties of the current securable object and get child securable objects with unique role assignments if any.
> ##### Parameters
> **obj:** The current securable object.

> **leafBreadthLimit:** Skip further visiting on this branch if the number of child items or documents with unique role assignments exceeded leafBreadthLimit.

> ##### Return value
> The child securable objects.

#### Visit(Microsoft.SharePoint.Client.SecurableObject,System.Int32,System.Action{Microsoft.SharePoint.Client.SecurableObject,System.String})
Traverse each descendents of a securable object with a specified action.
> ##### Parameters
> **obj:** The current securable object.

> **leafBreadthLimit:** Skip further visiting on this branch if the number of child items or documents with unique role assignments exceeded leafBreadthLimit.

> **action:** The action to be executed for each securable object.


#### GetUserEmail(Microsoft.SharePoint.Client.Web,System.Int32)
Get user email by user id.
> ##### Parameters
> **web:** The current web object.

> **userId:** The user id

> ##### Return value
> The email property of the specified user.

#### EnsureGroupCache(Microsoft.SharePoint.Client.SecurableObject,System.String)
Ensure all users of a given SharePoint group has been cached.
> ##### Parameters
> **obj:** The current securable object.

> **groupLoginName:** The group login name.


#### GetAllUniqueRoleAssignments(Microsoft.SharePoint.Client.Web,System.Int32)
Get all unique role assignments for a web object and all its descendents down to document or list item level.
> ##### Parameters
> **web:** The current web object to be processed.

> **leafBreadthLimit:** Skip further visiting on this branch if the number of child items or documents with unique role assignments exceeded leafBreadthLimit. When setting to 0, the process will stop at list / document library level.

> ##### Return value
> 

## SharePoint.Client.TaxonomyExtensions
            
Class for deprecated taxonomy extension methods
        
### Fields

#### TaxonomyGuidLabelDelimiter
The default Taxonomy Guid Label Delimiter
### Methods


#### CreateTermGroup(Microsoft.SharePoint.Client.Taxonomy.TermStore,System.String,System.Guid,System.String)
Creates a new term group, in the specified term store.
> ##### Parameters
> **termStore:** the term store to use

> **groupName:** Name of the term group

> **groupId:** (Optional) ID of the group; if not provided a random GUID is used

> **groupDescription:** (Optional) Description of the term group

> ##### Return value
> The created term group

#### EnsureTermGroup(Microsoft.SharePoint.Client.Site,System.String,System.Guid,System.String)
Ensures the named group exists, returning a reference to the group, and creating or updating as necessary.
> ##### Parameters
> **site:** Site connected to the term store to use

> **groupName:** Name of the term group

> **groupId:** (Optional) ID of the group; if not provided the parameter is ignored, a random GUID is used if necessary to create the group, otherwise if the ID differs a warning is logged

> **groupDescription:** (Optional) Description of the term group; if null or not provided the parameter is ignored, otherwise the group is updated as necessary to match the description; passing an empty string will clear the description

> ##### Return value
> The required term group

#### EnsureTermSet(Microsoft.SharePoint.Client.Taxonomy.TermGroup,System.String,System.Guid,System.Nullable{System.Int32},System.String,System.Nullable{System.Boolean},System.String,System.String)
Ensures the named term set exists, returning a reference to the set, and creating or updating as necessary.
> ##### Parameters
> **parentGroup:** Group to check or create the term set in

> **termSetName:** Name of the term set

> **termSetId:** (Optional) ID of the term set; if not provided the parameter is ignored, a random GUID is used if necessary to create the term set, otherwise if the ID differs a warning is logged

> **lcid:** (Optional) Default language of the term set; if not provided the default of the associate term store is used

> **description:** (Optional) Description of the term set; if null or not provided the parameter is ignored, otherwise the term set is updated as necessary to match the description; passing an empty string will clear the description

> **isOpen:** (Optional) Whether the term store is open for new term creation or not

> **termSetContact:** 

> **termSetOwner:** 

> ##### Return value
> The required term set

#### GetDefaultTermStore(Microsoft.SharePoint.Client.Web)
Private method used for resolving taxonomy term set for taxonomy field
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> ##### Return value
> 

#### GetTaxonomySession(Microsoft.SharePoint.Client.Site)
Returns a new taxonomy session for the current site
> ##### Parameters
> **site:** 

> ##### Return value
> 

#### GetDefaultKeywordsTermStore(Microsoft.SharePoint.Client.Site)
Returns the default keywords termstore for the current site
> ##### Parameters
> **site:** 

> ##### Return value
> 

#### GetDefaultSiteCollectionTermStore(Microsoft.SharePoint.Client.Site)
Returns the default site collection termstore
> ##### Parameters
> **site:** 

> ##### Return value
> 

#### GetTermSetsByName(Microsoft.SharePoint.Client.Site,System.String,System.Int32)
Finds a termset by name
> ##### Parameters
> **site:** The current site

> **name:** The name of the termset

> **lcid:** The locale ID for the termset to return, defaults to 1033

> ##### Return value
> 

#### GetTermGroupByName(Microsoft.SharePoint.Client.Site,System.String)
Finds a termgroup by name
> ##### Parameters
> **site:** The current site

> **name:** The name of the termgroup

> ##### Return value
> 

#### GetTermGroupByName(Microsoft.SharePoint.Client.Taxonomy.TermStore,System.String)
Gets the named term group, if it exists in the term store.
> ##### Parameters
> **termStore:** The term store to use

> **groupName:** Name of the term group

> ##### Return value
> The requested term group, or null if it does not exist

#### GetTermGroupById(Microsoft.SharePoint.Client.Site,System.Guid)
Finds a termgroup by its ID
> ##### Parameters
> **site:** The current site

> **termGroupId:** The ID of the termgroup

> ##### Return value
> 

#### GetTermByName(Microsoft.SharePoint.Client.Site,System.Guid,System.String)
Gets a Taxonomy Term by Name
> ##### Parameters
> **site:** The site to process

> **termSetId:** 

> **term:** 

> ##### Return value
> 

#### AddTermToTermset(Microsoft.SharePoint.Client.Site,System.Guid,System.String)
Adds a term to a given termset
> ##### Parameters
> **site:** The current site

> **termSetId:** The ID of the termset

> **term:** The label of the new term to create

> ##### Return value
> 

#### AddTermToTermset(Microsoft.SharePoint.Client.Site,System.Guid,System.String,System.Guid)
Adds a term to a given termset
> ##### Parameters
> **site:** The current site

> **termSetId:** The ID of the termset

> **term:** The label of the new term to create

> **termId:** The ID of the term to create

> ##### Return value
> 

#### ImportTerms(Microsoft.SharePoint.Client.Site,System.String[],System.Int32,System.String,System.Boolean)
Imports an array of | delimited strings into the deafult site collection termstore. Specify strings in this format: TermGroup|TermSet|Term E.g. "Locations|Nordics|Sweden"
> ##### Parameters
> **site:** 

> **termLines:** 

> **lcid:** 

> **delimiter:** 

> **synchronizeDeletions:** Remove tags that are not present in the import


#### ImportTerms(Microsoft.SharePoint.Client.Site,System.String[],System.Int32,Microsoft.SharePoint.Client.Taxonomy.TermStore,System.String,System.Boolean)
Imports an array of | delimited strings into the deafult site collection termstore. Specify strings in this format: TermGroup|TermSet|Term E.g. "Locations|Nordics|Sweden"
> ##### Parameters
> **site:** 

> **termLines:** 

> **lcid:** 

> **termStore:** The termstore to import the terms into

> **delimiter:** 

> **synchronizeDeletions:** Remove tags that are not present in the import


#### ImportTermSet(Microsoft.SharePoint.Client.Taxonomy.TermGroup,System.String,System.Guid,System.Boolean,System.Nullable{System.Boolean},System.String,System.String)
The format of the file is the same as that used by the import function in the web interface. A sample file can be obtained from the web interface. This is a CSV file, with the following headings: Term Set Name,Term Set Description,LCID,Available for Tagging,Term Description,Level 1 Term,Level 2 Term,Level 3 Term,Level 4 Term,Level 5 Term,Level 6 Term,Level 7 Term The first data row must contain the Term Set Name, Term Set Description, and LCID, and should also contain the first term. It is recommended that a fixed GUID be used as the termSetId, to allow the term set to be easily updated (so do not pass Guid.Empty). In contrast to the web interface import, this is not a one-off import but runs synchronisation logic allowing updating of an existing Term Set. When synchronising, any existing terms are matched (with Term Description and Available for Tagging updated as necessary), any new terms are added in the correct place in the hierarchy, and (if synchroniseDeletions is set) any terms not in the imported file are removed. The import file also supports an expanded syntax for the Term Set Name and term names (Level 1 Term, Level 2 Term, etc). These columns support values with the format "Name|GUID", with the name and GUID separated by a pipe character (note that the pipe character is invalid to use within a taxomony item name). This expanded syntax is not required, but can be used to ensure all terms have fixed IDs.
Imports terms from a term set file, updating with any new terms, in the same format at that used by the web interface import ability.
> ##### Parameters
> **termGroup:** Group to create the term set within

> **filePath:** Local path to the file to import

> **termSetId:** GUID to use for the term set; if Guid.Empty is passed then a random GUID is generated and used

> **synchroniseDeletions:** (Optional) Whether to also synchronise deletions; that is, remove any terms not in the import file; default is no (false)

> **termSetIsOpen:** (Optional) Whether the term set should be marked open; if not passed, then the existing setting is not changed

> **termSetContact:** (Optional) Contact for the term set; if not provided, the existing setting is retained

> **termSetOwner:** (Optional) Owner for the term set; if not provided, the existing setting is retained

> ##### Return value
> The created, or updated, term set

#### ImportTermSet(Microsoft.SharePoint.Client.Taxonomy.TermGroup,System.IO.Stream,System.Guid,System.Boolean,System.Nullable{System.Boolean},System.String,System.String)
The format of the file is the same as that used by the import function in the web interface. A sample file can be obtained from the web interface. This is a CSV file, with the following headings: Term Set Name,Term Set Description,LCID,Available for Tagging,Term Description,Level 1 Term,Level 2 Term,Level 3 Term,Level 4 Term,Level 5 Term,Level 6 Term,Level 7 Term The first data row must contain the Term Set Name, Term Set Description, and LCID, and should also contain the first term. It is recommended that a fixed GUID be used as the termSetId, to allow the term set to be easily updated (so do not pass Guid.Empty). In contrast to the web interface import, this is not a one-off import but runs synchronisation logic allowing updating of an existing Term Set. When synchronising, any existing terms are matched (with Term Description and Available for Tagging updated as necessary), any new terms are added in the correct place in the hierarchy, and (if synchroniseDeletions is set) any terms not in the imported file are removed. The import file also supports an expanded syntax for the Term Set Name and term names (Level 1 Term, Level 2 Term, etc). These columns support values with the format "Name|GUID", with the name and GUID separated by a pipe character (note that the pipe character is invalid to use within a taxomony item name). This expanded syntax is not required, but can be used to ensure all terms have fixed IDs.
Imports terms from a term set stream, updating with any new terms, in the same format at that used by the web interface import ability.
> ##### Parameters
> **termGroup:** Group to create the term set within

> **termSetData:** Stream containing the data to import

> **termSetId:** GUID to use for the term set; if Guid.Empty is passed then a random GUID is generated and used

> **synchroniseDeletions:** (Optional) Whether to also synchronise deletions; that is, remove any terms not in the import file; default is no (false)

> **termSetIsOpen:** (Optional) Whether the term set should be marked open; if not passed, then the existing setting is not changed

> **termSetContact:** (Optional) Contact for the term set; if not provided, the existing setting is retained

> **termSetOwner:** (Optional) Owner for the term set; if not provided, the existing setting is retained

> ##### Return value
> The created, or updated, term set

#### ExportTermSet(Microsoft.SharePoint.Client.Site,System.Guid,System.Boolean,System.String)
Exports the full list of terms from all termsets in all termstores.
> ##### Parameters
> **site:** The site to process

> **termSetId:** The ID of the termset to export

> **includeId:** if true, Ids of the the taxonomy items will be included

> **delimiter:** if specified, this delimiter will be used. Notice that IDs will be delimited with ;# from the label

> ##### Return value
> 

#### ExportTermSet(Microsoft.SharePoint.Client.Site,System.Guid,System.Boolean,Microsoft.SharePoint.Client.Taxonomy.TermStore,System.String)
Exports the full list of terms from all termsets in all termstores.
> ##### Parameters
> **site:** The site to export the termsets from

> **termSetId:** The ID of the termset to export

> **includeId:** if true, Ids of the the taxonomy items will be included

> **termStore:** The term store to export the termset from

> **delimiter:** if specified, this delimiter will be used. Notice that IDs will be delimited with ;# from the label

> ##### Return value
> 

#### ExportAllTerms(Microsoft.SharePoint.Client.Site,System.Boolean,System.String)
Exports the full list of terms from all termsets in all termstores.
> ##### Parameters
> **site:** The site to process

> **includeId:** if true, Ids of the the taxonomy items will be included

> **delimiter:** if specified, this delimiter will be used. Notice that IDs will be delimited with ;# from the label

> ##### Return value
> 

#### GetTaxonomyItemByPath(Microsoft.SharePoint.Client.Site,System.String,System.String)
Returns a taxonomy item by it's path, e.g. Group|Set|Term
> ##### Parameters
> **site:** The current site

> **path:** The path of the item to return

> **delimiter:** The delimeter separating groups, sets and term in the path. Defaults to |

> ##### Return value
> 

#### SetTaxonomyFieldValueByTermPath(Microsoft.SharePoint.Client.ListItem,System.String,System.Guid)
Sets a value in a taxonomy field
> ##### Parameters
> **item:** The item to set the value to

> **TermPath:** The path of the term in the shape of "TermGroupName|TermSetName|TermName"

> **fieldId:** The id of the field

> ##### Exceptions
> **System.Collections.Generic.KeyNotFoundException:** 


#### SetTaxonomyFieldValue(Microsoft.SharePoint.Client.ListItem,System.Guid,System.String,System.Guid)
Sets a value of a taxonomy field
> ##### Parameters
> **item:** The item to process

> **fieldId:** The ID of the field to set

> **label:** The label of the term to set

> **termGuid:** The id of the term to set


#### SetTaxonomyFieldValues(Microsoft.SharePoint.Client.ListItem,System.Guid,System.Collections.Generic.IEnumerable{System.Collections.Generic.KeyValuePair{System.Guid,System.String}})
Sets a value of a taxonomy field that supports multiple values
> ##### Parameters
> **item:** The item to process

> **fieldId:** The ID of the field to set

> **termValues:** The key and values of terms to set


#### CreateTaxonomyField(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.TaxonomyFieldCreationInformation)
Can be used to create taxonomy field remotely to web.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **fieldCreationInformation:** Creation Information of the field

> ##### Return value
> New taxonomy field

#### RemoveTaxonomyFieldByInternalName(Microsoft.SharePoint.Client.Web,System.String)
Removes a taxonomy field (site column) and its associated hidden field by internal name
> ##### Parameters
> **web:** Web object were the field (site column) exists

> **internalName:** Internal name of the taxonomy field (site column) to be removed


#### RemoveTaxonomyFieldById(Microsoft.SharePoint.Client.Web,System.Guid)
Removes a taxonomy field (site column) and its associated hidden field by id
> ##### Parameters
> **web:** Web object were the field (site column) exists

> **id:** Guid representing the id of the taxonomy field (site column) to be removed


#### CreateTaxonomyField(Microsoft.SharePoint.Client.List,OfficeDevPnP.Core.Entities.TaxonomyFieldCreationInformation)
Can be used to create taxonomy field remotely in a list.
> ##### Parameters
> **list:** List to be processed

> **fieldCreationInformation:** Creation information of the field

> ##### Return value
> New taxonomy field

#### WireUpTaxonomyField(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Field,System.String,System.String,System.Boolean)
Wires up MMS field to the specified term set.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **field:** Field to be wired up

> **mmsGroupName:** Taxonomy group

> **mmsTermSetName:** Term set name

> **multiValue:** If true, create a multivalue field


#### WireUpTaxonomyField(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Field,Microsoft.SharePoint.Client.Taxonomy.TermSet,System.Boolean)
Wires up MMS field to the specified term set.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **field:** Field to be wired up

> **termSet:** Taxonomy TermSet

> **multiValue:** If true, create a multivalue field


#### WireUpTaxonomyField(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Field,Microsoft.SharePoint.Client.Taxonomy.Term,System.Boolean)
Wires up MMS field to the specified term.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **field:** Field to be wired up

> **anchorTerm:** Taxonomy Term

> **multiValue:** If true, create a multivalue field


#### WireUpTaxonomyField(Microsoft.SharePoint.Client.Web,System.Guid,System.String,System.String,System.Boolean)
Wires up MMS field to the specified term set.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **id:** Field ID to be wired up

> **mmsGroupName:** Taxonomy group

> **mmsTermSetName:** Term set name

> **multiValue:** If true, create a multivalue field


#### WireUpTaxonomyField(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.Field,Microsoft.SharePoint.Client.Taxonomy.TermSet,System.Boolean)
Wires up MMS field to the specified term set.
> ##### Parameters
> **list:** List to be processed

> **field:** Field to be wired up

> **termSet:** Taxonomy TermSet

> **multiValue:** Term set name


#### WireUpTaxonomyField(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.Field,Microsoft.SharePoint.Client.Taxonomy.Term,System.Boolean)
Wires up MMS field to the specified term.
> ##### Parameters
> **list:** List to be processed

> **field:** Field to be wired up

> **anchorTerm:** Taxonomy Term

> **multiValue:** Allow multiple selection


#### WireUpTaxonomyField(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.Field,System.String,System.String,System.Boolean)
Wires up MMS field to the specified term set.
> ##### Parameters
> **list:** List to be processed

> **field:** Field to be wired up

> **mmsGroupName:** Taxonomy group

> **mmsTermSetName:** Term set name

> **multiValue:** Allow multiple selection


#### WireUpTaxonomyField(Microsoft.SharePoint.Client.List,System.Guid,System.String,System.String,System.Boolean)
Wires up MMS field to the specified term set.
> ##### Parameters
> **list:** List to be processed

> **id:** Field ID to be wired up

> **mmsGroupName:** Taxonomy group

> **mmsTermSetName:** Term set name

> **multiValue:** Allow multiple selection


#### WireUpTaxonomyFieldInternal(Microsoft.SharePoint.Client.Field,Microsoft.SharePoint.Client.Taxonomy.TaxonomyItem,System.Boolean)
Wires up MMS field to the specified term set or term.
> ##### Parameters
> **field:** Field to be wired up

> **taxonomyItem:** Taxonomy TermSet or Term

> **multiValue:** Allow multiple selection


#### GetWssIdForTerm(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.Taxonomy.Term)
Returns the Id for a term if present in the TaxonomyHiddenList. Otherwise returns -1;
> ##### Parameters
> **web:** 

> **term:** 

> ##### Return value
> 

#### SetTaxonomyFieldDefaultValue(Microsoft.SharePoint.Client.Field,Microsoft.SharePoint.Client.Taxonomy.TaxonomyItem,System.String,System.Boolean)
Sets the default value for a managed metadata field
> ##### Parameters
> **field:** Field to be wired up

> **taxonomyItem:** Taxonomy TermSet or Term

> **defaultValue:** default value for the field

> **pushChangesToLists:** push changes to lists


## SharePoint.Client.TenantExtensions
            
Class for deprecated tenant extension methods
        
### Methods


#### CreateSiteCollection(Microsoft.Online.SharePoint.TenantAdministration.Tenant,OfficeDevPnP.Core.Entities.SiteEntity,System.Boolean,System.Boolean,System.Func{OfficeDevPnP.Core.TenantOperationMessage,System.Boolean})
Adds a SiteEntity by launching site collection creation and waits for the creation to finish
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **properties:** Describes the site collection to be created

> **removeFromRecycleBin:** It true and site is present in recycle bin, it will be removed first from the recycle bin

> **wait:** If true, processing will halt until the site collection has been created

> **timeoutFunction:** An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.

> ##### Return value
> Guid of the created site collection and Guid.Empty is the wait parameter is specified as false. Returns Guid.Empty if the wait is cancelled.

#### CreateSiteCollection(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,System.String,System.String,System.String,System.Int32,System.Int32,System.Int32,System.Int32,System.Int32,System.UInt32,System.Boolean,System.Boolean,System.Func{OfficeDevPnP.Core.TenantOperationMessage,System.Boolean})
Launches a site collection creation and waits for the creation to finish
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** The SPO url

> **title:** The site title

> **siteOwnerLogin:** Owner account

> **template:** Site template being used

> **storageMaximumLevel:** Site quota in MB

> **storageWarningLevel:** Site quota warning level in MB

> **timeZoneId:** TimeZoneID for the site. "(UTC+01:00) Brussels, Copenhagen, Madrid, Paris" = 3

> **userCodeMaximumLevel:** The user code quota in points

> **userCodeWarningLevel:** The user code quota warning level in points

> **lcid:** The site locale. See http://technet.microsoft.com/en-us/library/ff463597.aspx for a complete list of Lcid's

> **removeFromRecycleBin:** If true, any existing site with the same URL will be removed from the recycle bin

> **wait:** Wait for the site to be created before continuing processing

> **timeoutFunction:** An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.

> ##### Return value
> 

#### CheckIfSiteExists(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,System.String)
Returns if a site collection is in a particular status. If the url contains a sub site then returns true is the sub site exists, false if not. Status is irrelevant for sub sites
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** Url to the site collection

> **status:** Status to check (Active, Creating, Recycled)

> ##### Return value
> True if in status, false if not in status

#### IsSiteActive(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String)
Checks if a site collection is Active
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** URL to the site collection

> ##### Return value
> True if active, false if not

#### SiteExists(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String)
Checks if a site collection exists, relies on tenant admin API. Sites that are recycled also return as existing sites
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** URL to the site collection

> ##### Return value
> True if existing, false if not

#### SubSiteExists(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String)
Checks if a sub site exists
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** URL to the sub site

> ##### Return value
> True if existing, false if not

#### DeleteSiteCollection(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,System.Boolean,System.Func{OfficeDevPnP.Core.TenantOperationMessage,System.Boolean})
Deletes a site collection
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** Url of the site collection to delete

> **useRecycleBin:** Leave the deleted site collection in the site collection recycle bin

> **timeoutFunction:** An optional function that will be called while waiting for the site to be created. Return true to cancel the wait loop.

> ##### Return value
> True if deleted

#### DeleteSiteCollectionFromRecycleBin(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,System.Boolean,System.Func{OfficeDevPnP.Core.TenantOperationMessage,System.Boolean})
Deletes a site collection from the site collection recycle bin
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** URL of the site collection to delete

> **wait:** If true, processing will halt until the site collection has been deleted from the recycle bin

> **timeoutFunction:** An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.


#### GetSiteGuidByUrl(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String)
Gets the ID of site collection with specified URL
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** A URL that specifies a site collection to get ID.

> ##### Return value
> The Guid of a site collection

#### GetSiteGuidByUrl(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.Uri)
Gets the ID of site collection with specified URL
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** A URL that specifies a site collection to get ID.

> ##### Return value
> The Guid of a site collection or an Guid.Empty if the Site does not exist

#### GetWebTemplates(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.UInt32,System.Int32)
Returns available webtemplates/site definitions
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **lcid:** 

> **compatibilityLevel:** 14 for SharePoint 2010, 15 for SharePoint 2013/SharePoint Online

> ##### Return value
> 

#### SetSiteProperties(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,System.String,System.Nullable{System.Boolean},System.Nullable{Microsoft.Online.SharePoint.TenantManagement.SharingCapabilities},System.Nullable{System.Int64},System.Nullable{System.Int64},System.Nullable{System.Double},System.Nullable{System.Double},System.Nullable{System.Boolean},System.Boolean,System.Func{OfficeDevPnP.Core.TenantOperationMessage,System.Boolean})
Sets tenant site Properties
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **siteFullUrl:** 

> **title:** 

> **allowSelfServiceUpgrade:** 

> **sharingCapability:** 

> **storageMaximumLevel:** 

> **storageWarningLevel:** 

> **userCodeMaximumLevel:** 

> **userCodeWarningLevel:** 

> **noScriptSite:** 


#### SetSiteLockState(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,OfficeDevPnP.Core.SiteLockState,System.Boolean,System.Func{OfficeDevPnP.Core.TenantOperationMessage,System.Boolean})
Sets a site to Unlock access or NoAccess. This operation may occur immediately, but the site lock may take a short while before it goes into effect.
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site (i.e. https://[tenant]-admin.sharepoint.com)

> **siteFullUrl:** The target site to change the lock state.

> **lockState:** The target state the site should be changed to.

> **wait:** If true, processing will halt until the site collection lock state has been implemented

> **timeoutFunction:** An optional function that will be called while waiting for the site to be created. If set will override the wait variable. Return true to cancel the wait loop.


#### AddAdministrators(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.Collections.Generic.IEnumerable{OfficeDevPnP.Core.Entities.UserEntity},System.Uri,System.Boolean)
Add a site collection administrator to a site collection
> ##### Parameters
> **tenant:** A tenant object pointing to the context of a Tenant Administration site

> **adminLogins:** Array of admins loginnames to add

> **siteUrl:** Url of the site to operate on

> **addToOwnersGroup:** Optionally the added admins can also be added to the Site owners group


#### GetSiteCollections(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.Int32,System.Int32,System.Boolean,System.Boolean)
Returns all site collections in the current Tenant based on a startIndex. IncludeDetail adds additional properties to the SPSite object.
> ##### Parameters
> **tenant:** Tenant object to operate against

> **startIndex:** Not relevant anymore

> **endIndex:** Not relevant anymore

> **includeDetail:** Option to return a limited set of data

> **includeOD4BSites:** Also return the OD4B sites

> ##### Return value
> An IList of SiteEntity objects

#### GetOneDriveSiteCollections(Microsoft.Online.SharePoint.TenantAdministration.Tenant)
Get OneDrive site collections by iterating through all user profiles.
> ##### Parameters
> **tenant:** 

> ##### Return value
> List of objects containing site collection info.

#### GetUserProfileServiceClient(Microsoft.Online.SharePoint.TenantAdministration.Tenant)
Gets the UserProfileService proxy to enable calls to the UPA web service.
> ##### Parameters
> **tenant:** 

> ##### Return value
> UserProfileService web service client

#### DeployApplicationPackageToAppCatalog(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,System.String,System.String,System.Boolean,System.Boolean)
Adds a package to the tenants app catalog and by default deploys it if the package is a client side package (sppkg)
> ##### Parameters
> **tenant:** Tenant to operate against

> **appCatalogSiteUrl:** Full url to the tenant admin site (e.g. https://contoso.sharepoint.com/sites/apps)

> **spPkgName:** Name of the package to upload (e.g. demo.sppkg)

> **spPkgPath:** Path on the filesystem where this package is stored

> **autoDeploy:** Automatically deploy the package, only applies to client side packages (sppkg)

> **overwrite:** Overwrite the package if it was already listed in the app catalog

> ##### Return value
> The ListItem of the added package row

## SharePoint.Client.WebExtensions
            
Class that holds deprecated methods for site (both site collection and web site) creation, status, retrieval and settings
            
Class that deals with site (both site collection and web site) creation, status, retrieval and settings
        
### Methods


#### GetBaseTemplateId(Microsoft.SharePoint.Client.Web)
Returns the Base Template ID for the current web
> ##### Parameters
> **parentWeb:** The parent Web (site) to get the base template from

> ##### Return value
> The Base Template ID for the current web

#### CreateWeb(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Entities.SiteEntity,System.Boolean,System.Boolean)
Adds a new child Web (site) to a parent Web.
> ##### Parameters
> **parentWeb:** The parent Web (site) to create under

> **subsite:** Details of the Web (site) to add. Only Title, Url (as the leaf URL), Description, Template and Language are used.

> **inheritPermissions:** Specifies whether the new site will inherit permissions from its parent site.

> **inheritNavigation:** Specifies whether the site inherits navigation.

> ##### Return value
> 

#### CreateWeb(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String,System.String,System.Int32,System.Boolean,System.Boolean)
Adds a new child Web (site) to a parent Web.
> ##### Parameters
> **parentWeb:** The parent Web (site) to create under

> **title:** The title of the new site.

> **leafUrl:** A string that represents the URL leaf name.

> **description:** The description of the new site.

> **template:** The name of the site template to be used for creating the new site.

> **language:** The locale ID that specifies the language of the new site.

> **inheritPermissions:** Specifies whether the new site will inherit permissions from its parent site.

> **inheritNavigation:** Specifies whether the site inherits navigation.


#### DeleteWeb(Microsoft.SharePoint.Client.Web,System.String)
Deletes the child website with the specified leaf URL, from a parent Web, if it exists.
> ##### Parameters
> **parentWeb:** The parent Web (site) to delete from

> **leafUrl:** A string that represents the URL leaf name.

> ##### Return value
> true if the web was deleted; otherwise false if nothing was done

#### GetAllWebUrls(Microsoft.SharePoint.Client.Site)
This is analagous to the SPSite.AllWebs property and can be used to get a collection of all web site URLs to loop through, e.g. for branding.
Gets the collection of the URLs of all Web sites that are contained within the site collection, including the top-level site and its subsites.
> ##### Parameters
> **site:** Site collection to retrieve the URLs for.

> ##### Return value
> An enumeration containing the full URLs as strings.

#### GetWeb(Microsoft.SharePoint.Client.Web,System.String)
The ServerRelativeUrl property of the retrieved Web is instantiated.
Returns the child Web site with the specified leaf URL.
> ##### Parameters
> **parentWeb:** The Web site to check under

> **leafUrl:** A string that represents the URL leaf name.

> ##### Return value
> The requested Web, if it exists, otherwise null.

#### WebExists(Microsoft.SharePoint.Client.Web,System.String)
Determines if a child Web site with the specified leaf URL exists.
> ##### Parameters
> **parentWeb:** The Web site to check under

> **leafUrl:** A string that represents the URL leaf name.

> ##### Return value
> true if the Web (site) exists; otherwise false

#### WebExistsFullUrl(Microsoft.SharePoint.Client.ClientRuntimeContext,System.String)
Determines if a Web (site) exists at the specified full URL, either accessible or that returns an access error.
> ##### Parameters
> **context:** Existing context, used to provide credentials.

> **webFullUrl:** Full URL of the site to check.

> ##### Return value
> true if the Web (site) exists; otherwise false

#### WebExistsByTitle(Microsoft.SharePoint.Client.Web,System.String)
Determines if a web exists by title.
> ##### Parameters
> **title:** Title of the web to check.

> **parentWeb:** Parent web to check under.

> ##### Return value
> True if a web with the given title exists.

#### IsSubSite(Microsoft.SharePoint.Client.Web)
Checks if the current web is a sub site or not
> ##### Parameters
> **web:** Web to check

> ##### Return value
> True is sub site, false otherwise

#### IsNoScriptSite(Microsoft.SharePoint.Client.Site)
Detects if the site in question has no script enabled or not. Detection is done by verifying if the AddAndCustomizePages permission is missing. See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f for the effects of NoScript
> ##### Parameters
> **site:** site to verify

> ##### Return value
> True if noscript, false otherwise

#### IsNoScriptSite(Microsoft.SharePoint.Client.Web)
Detects if the site in question has no script enabled or not. Detection is done by verifying if the AddAndCustomizePages permission is missing. See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f for the effects of NoScript
> ##### Parameters
> **web:** Web to verify

> ##### Return value
> True if noscript, false otherwise

#### GetAppInstances(Microsoft.SharePoint.Client.Web,System.Linq.Expressions.Expression{System.Func{Microsoft.SharePoint.Client.AppInstance,System.Object}}[])
Returns all app instances
> ##### Parameters
> **web:** The site to process

> **expressions:** List of lambda expressions of properties to load when retrieving the object

> ##### Return value
> 

#### RemoveAppInstanceByTitle(Microsoft.SharePoint.Client.Web,System.String)
Removes the app instance with the specified title.
> ##### Parameters
> **web:** Web to remove the app instance from

> **appTitle:** Title of the app instance to remove

> ##### Return value
> true if the the app instance was removed; false if it does not exist

#### InstallSolution(Microsoft.SharePoint.Client.Site,System.Guid,System.String,System.Int32,System.Int32)
Uploads and installs a sandbox solution package (.WSP) file, replacing existing solution if necessary.
> ##### Parameters
> **site:** Site collection to install to

> **packageGuid:** ID of the solution, from the solution manifest (required for the remove step)

> **sourceFilePath:** Path to the sandbox solution package (.WSP) file

> **majorVersion:** Optional major version of the solution, defaults to 1

> **minorVersion:** Optional minor version of the solution, defaults to 0


#### UninstallSolution(Microsoft.SharePoint.Client.Site,System.Guid,System.String,System.Int32,System.Int32)
Uninstalls a sandbox solution package (.WSP) file
> ##### Parameters
> **site:** Site collection to install to

> **packageGuid:** ID of the solution, from the solution manifest

> **fileName:** filename of the WSP file to uninstall

> **majorVersion:** Optional major version of the solution, defaults to 1

> **minorVersion:** Optional minor version of the solution, defaults to 0


#### MySiteSearch(Microsoft.SharePoint.Client.Web)
Returns all my site site collections
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> ##### Return value
> All my site site collections

#### SiteSearch(Microsoft.SharePoint.Client.Web)
Returns all site collections that are indexed. In MT the search center, mysite host and contenttype hub are defined as non indexable by default and thus are not returned
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> ##### Return value
> All site collections

#### SiteSearch(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Returns the site collections that comply with the passed keyword query
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **keywordQueryValue:** Keyword query

> **trimDuplicates:** Indicates if dublicates should be trimmed or not

> ##### Return value
> All found site collections

#### SiteSearchScopedByUrl(Microsoft.SharePoint.Client.Web,System.String)
Returns all site collection that start with the provided URL
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **siteUrl:** Base URL for which sites can be returned

> ##### Return value
> All found site collections

#### SiteSearchScopedByTitle(Microsoft.SharePoint.Client.Web,System.String)
Returns all site collection that match with the provided title
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **siteTitle:** Title of the site to search for

> ##### Return value
> All found site collections

#### ProcessQuery(Microsoft.SharePoint.Client.Web,System.String,System.Collections.Generic.List{OfficeDevPnP.Core.Entities.SiteEntity},Microsoft.SharePoint.Client.Search.Query.KeywordQuery)
Runs a query
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **keywordQueryValue:** keyword query

> **sites:** sites variable that hold the resulting sites

> **keywordQuery:** KeywordQuery object

> ##### Return value
> Total number of rows for the query

#### SetPropertyBagValue(Microsoft.SharePoint.Client.Web,System.String,System.Int32)
Sets a key/value pair in the web property bag
> ##### Parameters
> **web:** Web that will hold the property bag entry

> **key:** Key for the property bag entry

> **value:** Integer value for the property bag entry


#### SetPropertyBagValue(Microsoft.SharePoint.Client.Web,System.String,System.String)
Sets a key/value pair in the web property bag
> ##### Parameters
> **web:** Web that will hold the property bag entry

> **key:** Key for the property bag entry

> **value:** String value for the property bag entry


#### SetPropertyBagValue(Microsoft.SharePoint.Client.Web,System.String,System.DateTime)
Sets a key/value pair in the web property bag
> ##### Parameters
> **web:** Web that will hold the property bag entry

> **key:** Key for the property bag entry

> **value:** Datetime value for the property bag entry


#### SetPropertyBagValueInternal(Microsoft.SharePoint.Client.Web,System.String,System.Object)
Sets a key/value pair in the web property bag
> ##### Parameters
> **web:** Web that will hold the property bag entry

> **key:** Key for the property bag entry

> **value:** Value for the property bag entry


#### RemovePropertyBagValue(Microsoft.SharePoint.Client.Web,System.String)
Removes a property bag value from the property bag
> ##### Parameters
> **web:** The site to process

> **key:** The key to remove


#### RemovePropertyBagValueInternal(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Removes a property bag value
> ##### Parameters
> **web:** The web to process

> **key:** They key to remove

> **checkIndexed:** 


#### GetPropertyBagValueInt(Microsoft.SharePoint.Client.Web,System.String,System.Int32)
Get int typed property bag value. If does not contain, returns default value.
> ##### Parameters
> **web:** Web to read the property bag value from

> **key:** Key of the property bag entry to return

> **defaultValue:** 

> ##### Return value
> Value of the property bag entry as integer

#### GetPropertyBagValueDateTime(Microsoft.SharePoint.Client.Web,System.String,System.DateTime)
Get DateTime typed property bag value. If does not contain, returns default value.
> ##### Parameters
> **web:** Web to read the property bag value from

> **key:** Key of the property bag entry to return

> **defaultValue:** 

> ##### Return value
> Value of the property bag entry as integer

#### GetPropertyBagValueString(Microsoft.SharePoint.Client.Web,System.String,System.String)
Get string typed property bag value. If does not contain, returns given default value.
> ##### Parameters
> **web:** Web to read the property bag value from

> **key:** Key of the property bag entry to return

> **defaultValue:** 

> ##### Return value
> Value of the property bag entry as string

#### GetPropertyBagValueInternal(Microsoft.SharePoint.Client.Web,System.String)
Type independent implementation of the property getter.
> ##### Parameters
> **web:** Web to read the property bag value from

> **key:** Key of the property bag entry to return

> ##### Return value
> Value of the property bag entry

#### PropertyBagContainsKey(Microsoft.SharePoint.Client.Web,System.String)
Checks if the given property bag entry exists
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **key:** Key of the property bag entry to check

> ##### Return value
> True if the entry exists, false otherwise

#### GetEncodedValueForSearchIndexProperty(System.Collections.Generic.IEnumerable{System.String})
Used to convert the list of property keys is required format for listing keys to be index
> ##### Parameters
> **keys:** list of keys to set to be searchable

> ##### Return value
> string formatted list of keys in proper format

#### GetIndexedPropertyBagKeys(Microsoft.SharePoint.Client.Web)
Returns all keys in the property bag that have been marked for indexing
> ##### Parameters
> **web:** The site to process

> ##### Return value
> 

#### AddIndexedPropertyBagKey(Microsoft.SharePoint.Client.Web,System.String)
Marks a property bag key for indexing
> ##### Parameters
> **web:** The web to process

> **key:** The key to mark for indexing

> ##### Return value
> Returns True if succeeded

#### RemoveIndexedPropertyBagKey(Microsoft.SharePoint.Client.Web,System.String)
Unmarks a property bag key for indexing
> ##### Parameters
> **web:** The site to process

> **key:** The key to unmark for indexed. Case-sensitive

> ##### Return value
> Returns True if succeeded

#### ReIndexWeb(Microsoft.SharePoint.Client.Web)
Queues a web for a full crawl the next incremental/continous crawl
> ##### Parameters
> **web:** Site to be processed


#### AddRemoteEventReceiver(Microsoft.SharePoint.Client.Web,System.String,System.String,Microsoft.SharePoint.Client.EventReceiverType,Microsoft.SharePoint.Client.EventReceiverSynchronization,System.Boolean)
Registers a remote event receiver
> ##### Parameters
> **web:** The web to process

> **name:** The name of the event receiver (needs to be unique among the event receivers registered on this list)

> **url:** The URL of the remote WCF service that handles the event

> **eventReceiverType:** 

> **synchronization:** 

> **force:** If True any event already registered with the same name will be removed first.

> ##### Return value
> Returns an EventReceiverDefinition if succeeded. Returns null if failed.

#### AddRemoteEventReceiver(Microsoft.SharePoint.Client.Web,System.String,System.String,Microsoft.SharePoint.Client.EventReceiverType,Microsoft.SharePoint.Client.EventReceiverSynchronization,System.Int32,System.Boolean)
Registers a remote event receiver
> ##### Parameters
> **web:** The web to process

> **name:** The name of the event receiver (needs to be unique among the event receivers registered on this list)

> **url:** The URL of the remote WCF service that handles the event

> **eventReceiverType:** 

> **synchronization:** 

> **sequenceNumber:** 

> **force:** If True any event already registered with the same name will be removed first.

> ##### Return value
> Returns an EventReceiverDefinition if succeeded. Returns null if failed.

#### GetEventReceiverById(Microsoft.SharePoint.Client.Web,System.Guid)
Returns an event receiver definition
> ##### Parameters
> **web:** Web to process

> **id:** 

> ##### Return value
> 

#### GetEventReceiverByName(Microsoft.SharePoint.Client.Web,System.String)
Returns an event receiver definition
> ##### Parameters
> **web:** 

> **name:** 

> ##### Return value
> 

#### SetLocalizationLabels(Microsoft.SharePoint.Client.Web,System.String,System.String,System.String)
Can be used to set translations for different cultures.
> ##### Parameters
> **web:** Site to be processed - can be root web or sub site

> **cultureName:** Culture name like en-us or fi-fi

> **titleResource:** Localized Title string

> **descriptionResource:** Localized Description string

> ##### Example
> 
                web.SetLocalizationForSiteLabels("fi-fi", "Name of the site in Finnish", "Description in Finnish");
            

#### ApplyProvisioningTemplate(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.ProvisioningTemplateApplyingInformation)
Can be used to apply custom remote provisioning template on top of existing site.
> ##### Parameters
> **web:** 

> **template:** ProvisioningTemplate with the settings to be applied

> **applyingInformation:** Specified additional settings and or properties


#### GetProvisioningTemplate(Microsoft.SharePoint.Client.Web)
Can be used to extract custom provisioning template from existing site. The extracted template will be compared with the default base template.
> ##### Parameters
> **web:** Web to get template from

> ##### Return value
> ProvisioningTemplate object with generated values from existing site

#### GetProvisioningTemplate(Microsoft.SharePoint.Client.Web,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.ProvisioningTemplateCreationInformation)
Can be used to extract custom provisioning template from existing site. The extracted template will be compared with the default base template.
> ##### Parameters
> **web:** Web to get template from

> **creationInfo:** Specifies additional settings and/or properties

> ##### Return value
> ProvisioningTemplate object with generated values from existing site

#### SetPageOutputCache(Microsoft.SharePoint.Client.Web,System.Boolean,System.Int32,System.Int32,System.Boolean)
Sets output cache on publishing web. The settings can be maintained from UI by visiting url /_layouts/15/sitecachesettings.aspx
> ##### Parameters
> **web:** SharePoint web

> **enableOutputCache:** Specify true to enable output cache. False otherwise.

> **anonymousCacheProfileId:** Applies for anonymous users access for a site in Site Collection. Id of the profile specified in "Cache Profiles" list.

> **authenticatedCacheProfileId:** Applies for authenticated users access for a site in the Site Collection. Id of the profile specified in "Cache Profiles" list.

> **debugCacheInformation:** Specify true to enable the display of additional cache information on pages in this site collection. False otherwise.


#### DisableRequestAccess(Microsoft.SharePoint.Client.Web)
Disables the request access on the web.
> ##### Parameters
> **web:** The web to disable request access.


#### EnableRequestAccess(Microsoft.SharePoint.Client.Web,System.String[])
Enables request access for the specified e-mail addresses.
> ##### Parameters
> **web:** The web to enable request access.

> **emails:** The e-mail addresses to send access requests to.


#### EnableRequestAccess(Microsoft.SharePoint.Client.Web,System.Collections.Generic.IEnumerable{System.String})
Enables request access for the specified e-mail addresses.
> ##### Parameters
> **web:** The web to enable request access.

> **emails:** The e-mail addresses to send access requests to.


#### GetRequestAccessEmails(Microsoft.SharePoint.Client.Web)
Gets the request access e-mail addresses of the web.
> ##### Parameters
> **web:** The web to get the request access e-mail addresses from.

> ##### Return value
> The request access e-mail addresses of the web.

## SharePoint.Client.WorkflowExtensions
            
Class for deprecated workflow extension methods
        
### Methods


#### GetWorkflowSubscription(Microsoft.SharePoint.Client.Web,System.String)
Returns a workflow subscription for a site.
> ##### Parameters
> **web:** 

> **name:** 

> ##### Return value
> 

#### GetWorkflowSubscription(Microsoft.SharePoint.Client.Web,System.Guid)
Returns a workflow subscription
> ##### Parameters
> **web:** 

> **id:** 

> ##### Return value
> 

#### GetWorkflowSubscription(Microsoft.SharePoint.Client.List,System.String)
Returns a workflow subscription (associations) for a list
> ##### Parameters
> **list:** 

> **name:** 

> ##### Return value
> 

#### GetWorkflowSubscriptions(Microsoft.SharePoint.Client.Web)
Returns all the workflow subscriptions (associations) for the web and the lists of that web
> ##### Parameters
> **web:** The target Web

> ##### Return value
> 

#### AddWorkflowSubscription(Microsoft.SharePoint.Client.List,System.String,System.String,System.Boolean,System.Boolean,System.Boolean,System.String,System.String,System.Collections.Generic.Dictionary{System.String,System.String})
Adds a workflow subscription
> ##### Parameters
> **list:** 

> **workflowDefinitionName:** The name of the workflow definition WorkflowExtensions.GetWorkflowDefinition

> **subscriptionName:** The name of the workflow subscription to create

> **startManually:** if True the workflow can be started manually

> **startOnCreate:** if True the workflow will be started on item creation

> **startOnChange:** if True the workflow will be started on item change

> **historyListName:** the name of the history list. If not available it will be created

> **taskListName:** the name of the task list. If not available it will be created

> **associationValues:** 

> ##### Return value
> Guid of the workflow subscription

#### AddWorkflowSubscription(Microsoft.SharePoint.Client.List,Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition,System.String,System.Boolean,System.Boolean,System.Boolean,System.String,System.String,System.Collections.Generic.Dictionary{System.String,System.String})
Adds a workflow subscription to a list
> ##### Parameters
> **list:** 

> **workflowDefinition:** The workflow definition. WorkflowExtensions.GetWorkflowDefinition

> **subscriptionName:** The name of the workflow subscription to create

> **startManually:** if True the workflow can be started manually

> **startOnCreate:** if True the workflow will be started on item creation

> **startOnChange:** if True the workflow will be started on item change

> **historyListName:** the name of the history list. If not available it will be created

> **taskListName:** the name of the task list. If not available it will be created

> **associationValues:** 

> ##### Return value
> Guid of the workflow subscription

#### Delete(Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription)
Deletes the subscription
> ##### Parameters
> **subscription:** 


#### GetWorkflowDefinition(Microsoft.SharePoint.Client.Web,System.String,System.Boolean)
Returns a workflow definition for a site
> ##### Parameters
> **web:** 

> **displayName:** 

> **publishedOnly:** 

> ##### Return value
> 

#### GetWorkflowDefinition(Microsoft.SharePoint.Client.Web,System.Guid)
Returns a workflow definition
> ##### Parameters
> **web:** 

> **id:** 

> ##### Return value
> 

#### GetWorkflowDefinitions(Microsoft.SharePoint.Client.Web,System.Boolean)
Returns all the workflow definitions
> ##### Parameters
> **web:** The target Web

> **publishedOnly:** Defines whether to include only published definition, or all the definitions

> ##### Return value
> 

#### Delete(Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition)
Deletes a workflow definition
> ##### Parameters
> **definition:** 


#### GetWorkflowInstances(Microsoft.SharePoint.Client.Web)
Returns alls workflow instances for a site
> ##### Parameters
> **web:** 

> ##### Return value
> 

#### GetWorkflowInstances(Microsoft.SharePoint.Client.Web,Microsoft.SharePoint.Client.ListItem)
Returns alls workflow instances for a list item
> ##### Parameters
> **web:** 

> **item:** 

> ##### Return value
> 

#### GetInstances(Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription)
Returns all instances of a workflow for this subscription
> ##### Parameters
> **subscription:** 

> ##### Return value
> 

#### CancelWorkFlow(Microsoft.SharePoint.Client.WorkflowServices.WorkflowInstance)
Cancels a workflow instance
> ##### Parameters
> **instance:** 


#### ResumeWorkflow(Microsoft.SharePoint.Client.WorkflowServices.WorkflowInstance)
Resumes a workflow
> ##### Parameters
> **instance:** 


#### PublishCustomEvent(Microsoft.SharePoint.Client.WorkflowServices.WorkflowInstance,System.String,System.String)
Publish a custom event to a target workflow instance
> ##### Parameters
> **instance:** 

> **eventName:** The name of the target event

> **payload:** The payload that will be sent to the event


#### StartWorkflowInstance(Microsoft.SharePoint.Client.Web,System.String,System.Collections.Generic.IDictionary{System.String,System.Object})
Starts a new instance of a workflow definition against the current web site
> ##### Parameters
> **web:** The target web site

> **subscriptionName:** The name of the workflow subscription to start

> **payload:** Any input argument for the workflow instance

> ##### Return value
> The ID of the just started workflow instance

#### StartWorkflowInstance(Microsoft.SharePoint.Client.Web,System.Guid,System.Collections.Generic.IDictionary{System.String,System.Object})
Starts a new instance of a workflow definition against the current web site
> ##### Parameters
> **web:** The target web site

> **subscriptionId:** The ID of the workflow subscription to start

> **payload:** Any input argument for the workflow instance

> ##### Return value
> The ID of the just started workflow instance

#### StartWorkflowInstance(Microsoft.SharePoint.Client.ListItem,System.String,System.Collections.Generic.IDictionary{System.String,System.Object})
Starts a new instance of a workflow definition against the current item
> ##### Parameters
> **item:** The target item

> **subscriptionName:** The name of the workflow subscription to start

> **payload:** Any input argument for the workflow instance

> ##### Return value
> The ID of the just started workflow instance

#### StartWorkflowInstance(Microsoft.SharePoint.Client.ListItem,System.Guid,System.Collections.Generic.IDictionary{System.String,System.Object})
Starts a new instance of a workflow definition against the current item
> ##### Parameters
> **item:** The target item

> **subscriptionId:** The ID of the workflow subscription to start

> **payload:** Any input argument for the workflow instance

> ##### Return value
> The ID of the just started workflow instance

## SharePoint.Client.ListRatingExtensions
            
Enables: Ratings / Likes functionality on list in publishing web.
        
### Fields

#### RatingsFieldGuid_AverageRating
TODO: Replace Logging throughout
### Methods


#### SetRating(Microsoft.SharePoint.Client.List,OfficeDevPnP.Core.VotingExperience)
Enable Social Settings Likes/Ratings on list. Note: 1. Requires Publishing feature enabled on the web. 2. Defaults enable Ratings Experience on the List 3. When experience set to None, fields are not removed from the list since CSOM does not support removing hidden fields
> ##### Parameters
> **list:** Current List

> **experience:** Likes/Ratings


#### RemoveViewFields
Removes the view fields associated with any rating type

#### AddListFields
Add Ratings/Likes related fields to List from current Web

#### AddViewFields(OfficeDevPnP.Core.VotingExperience)
Add/Remove Ratings/Likes field in default view depending on exerpeince selected
> ##### Parameters
> **experience:** 


#### RemoveField(Microsoft.SharePoint.Client.List,System.Guid)
Removes a field from a list
> ##### Parameters
> **list:** 

> **fieldId:** 


#### EnsureField(Microsoft.SharePoint.Client.List,System.Guid)
Check for Ratings/Likes field and add to ListField if doesn't exists.
> ##### Parameters
> **list:** List

> **fieldId:** Field Id

> ##### Return value
> 

#### SetProperty(OfficeDevPnP.Core.VotingExperience)
Add required key/value settings on List Root-Folder
> ##### Parameters
> **experience:** 


## SharePoint.Client.VariationExtensions
            
Class that provides methods for variations
        
### Methods


#### ConfigureVariationsSettings(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Entities.VariationInformation)
Configures the variation settings 1. Go to "Site Actions" -> "Site settings" 2. Under "Site collection administration", click "Variation Settings". This method is for the page above to change or update the "Variation Settings"
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **variationSettings:** Variation settings


#### ProvisionSourceVariationLabel(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Entities.VariationLabelEntity)
Creates source variation label
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **sourceVariationLabel:** Source variation label


#### ProvisionTargetVariationLabels(Microsoft.SharePoint.Client.ClientContext,System.Collections.Generic.List{OfficeDevPnP.Core.Entities.VariationLabelEntity})
Creates target variation labels
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **variationLabels:** Variation labels


#### WaitForVariationLabelCreation(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Entities.VariationLabelEntity)
Wait for the variation label creation
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **variationLabel:** Variation label


#### GetVariationLabels(Microsoft.SharePoint.Client.ClientContext)
Retrieve all configured variation labels
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> ##### Return value
> Collection of VariationLabelEntity objects

#### CreateVariationLabels(Microsoft.SharePoint.Client.ClientContext,System.Collections.Generic.List{OfficeDevPnP.Core.Entities.VariationLabelEntity})
Create variation labels
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **variationLabels:** Variation labels


#### CheckForHierarchyCreation(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Entities.VariationLabelEntity)
Checks if hierarchy is created for the variation label. Get the "Hierarchy_x0020_Is_x0020_Created" list item value
> ##### Parameters
> **context:** Context for SharePoint objects and operations

> **variationLabel:** Variation label

> ##### Return value
> True, if hierarchy is created for the variation label

## SharePoint.Client.BaseTemplateManager
            
This class will be used to provide access to the right base template configuration
        

## SharePoint.Client.ManagedNavigationKind
            
Defines the kind of Managed Navigation for a site
        
### Fields

#### Current
Current Navigation
#### Global
Global Navigation

## Core.AppModelExtensions.VariationExtensions
            
Class that holds deprecated methods for variations
        

## Core.CoreResources
            
A strongly-typed resource class, for looking up localized strings, etc.
        
### Properties

#### ResourceManager
Returns the cached ResourceManager instance used by this class.
#### Culture
Overrides the current thread's CurrentUICulture property for all resource lookups using this strongly typed resource class.
#### AuthenticationManager_GetContext
Looks up a localized string similar to Getting authentication context for '{0}'.
#### AuthenticationManager_TenantUser
Looks up a localized string similar to Tenant user '{0}'.
#### AuthenticationManger_ProblemDeterminingTokenLease
Looks up a localized string similar to Could not determine lease for appOnlyAccessToken. Error = {0}.
#### BrandingExtension_ApplyTheme
Looks up a localized string similar to Applying theme '{0}' in '{1}'.
#### BrandingExtension_ComposedLookMissing
Looks up a localized string similar to Composed look '{0}' not found..
#### BrandingExtension_CreateComposedLook
Looks up a localized string similar to Creating composed look '{0}' in '{1}'.
#### BrandingExtension_DeployMasterPage
Looks up a localized string similar to Deploying masterpage '{0}' to '{1}'..
#### BrandingExtension_DeployPageLayout
Looks up a localized string similar to Deploying page layout '{0}' to '{1}'..
#### BrandingExtension_DeployTheme
Looks up a localized string similar to Deploying theme '{0}' to '{1}'.
#### BrandingExtension_InvalidPageLayoutName
Looks up a localized string similar to Cannot find Page Layout with name '{0}'..
#### BrandingExtension_SetCustomMasterUrl
Looks up a localized string similar to Setting custom master URL '{0}' in '{1}'..
#### BrandingExtension_SetMasterUrl
Looks up a localized string similar to Setting master URL '{0}' in '{1}'..
#### BrandingExtension_UpdateComposedLook
Looks up a localized string similar to Updating composed look '{0}' in '{1}'.
#### BrandingExtensions_UploadThemeFile_Destination_file_name_is_required_
Looks up a localized string similar to Destination file name is required..
#### BrandingExtensions_UploadThemeFile_Source_file_path_is_required_
Looks up a localized string similar to Source file path is required..
#### BrandingExtensions_UploadThemeFile_The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_
Looks up a localized string similar to The argument must be a single file name and cannot contain path characters..
#### ClientContextExtensions_Clone_Url_of_the_site_is_required_
Looks up a localized string similar to Url of the site is required..
#### ClientContextExtensions_ExecuteQueryRetry
Looks up a localized string similar to CSOM request frequency exceeded usage limits. Sleeping for {0} milliseconds before retrying..
#### ClientContextExtensions_ExecuteQueryRetryException
Looks up a localized string similar to ExecuteQuery threw following exception: {0}..
#### ClientContextExtensions_HasMinimalServerLibraryVersion_Error
Looks up a localized string similar to The server version could not be detected. Note that the check does assume the process at least has read access to SharePoint. Error: {0}..
#### Exception_Message_EmptyString_Arg
Looks up a localized string similar to The passed argument is a zero-length string or contains only whitespace..
#### FeatureExtensions_ActivateSiteCollectionFeature
Looks up a localized string similar to Activating feature {0} in site collection..
#### FeatureExtensions_ActivateWebFeature
Looks up a localized string similar to Activating feature {0} in web..
#### FeatureExtensions_DeactivateSiteCollectionFeature
Looks up a localized string similar to Deactivating feature {0} in site collection..
#### FeatureExtensions_DeactivateWebFeature
Looks up a localized string similar to Deactivating feature {0} in web..
#### FeatureExtensions_FeatureActivationProblem
Looks up a localized string similar to Problem with activation for feature id {0}. Error = {1}.
#### FeatureExtensions_ProcessFeatureInternal_FeatureActivationState
Looks up a localized string similar to Activation state for feature with id {1} was {0}..
#### FeatureExtensions_ProcessFeatureInternal_FeatureActive
Looks up a localized string similar to Feature activation for {0} returned success..
#### FeatureExtensions_ProcessFeatureInternal_FeatureException
Looks up a localized string similar to Error caught while waiting for ExecuteQueryRetry to complete. Error = {0}..
#### FieldAndContentTypeExtensions_AddField0ToContentType1
Looks up a localized string similar to Adding field ({0}) to content type ({1})..
#### FieldAndContentTypeExtensions_ContentType01AlreadyExists
Looks up a localized string similar to Content type '{0}' ({1}) already exists; no changes made..
#### FieldAndContentTypeExtensions_CreateContentType01
Looks up a localized string similar to Creating content type '{0}' ({1})..
#### FieldAndContentTypeExtensions_CreateDocumentSet
Looks up a localized string similar to Creating document set '{0}'..
#### FieldAndContentTypeExtensions_CreateField01
Looks up a localized string similar to Creating field '{0}' ({1})..
#### FieldAndContentTypeExtensions_CreateFieldBase
Looks up a localized string similar to New Field as XML: {0}.
#### FieldAndContentTypeExtensions_DeleteContentTypeById
Looks up a localized string similar to Could not find content type with id: {0}.
#### FieldAndContentTypeExtensions_DeleteContentTypeByName
Looks up a localized string similar to Could not find content type with name: {0}.
#### FieldAndContentTypeExtensions_Field01AlreadyExists
Looks up a localized string similar to Field '{0}' ({1}) already exists; no changes made..
#### FileFolderExtensions_CreateDocumentSet_The_argument_must_be_a_single_document_set_name_and_cannot_contain_path_characters_
Looks up a localized string similar to The argument must be a single document set name and cannot contain path characters..
#### FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_
Looks up a localized string similar to The argument must be a single folder name and cannot contain path characters..
#### FileFolderExtensions_CreateFolder0Under12
Looks up a localized string similar to Creating folder '{0}' under {1} '{2}'..
#### FileFolderExtensions_EnsureFolderPath_Folder_URL_is_required_
Looks up a localized string similar to Folder URL is required..
#### FileFolderExtensions_FolderMissing
Looks up a localized string similar to Target folder does not exist in the web. Web: {0}, Folder: {1}.
#### FileFolderExtensions_LibraryMissing
Looks up a localized string similar to Target library does not exist in the web. Web: {0}, List: {1}.
#### FileFolderExtensions_SetFileProperties_Error
Looks up a localized string similar to Content Type {0} does not exist in target list!.
#### FileFolderExtensions_UpdateFile0Properties1
Looks up a localized string similar to Update file '{0}', change properties: {1}..
#### FileFolderExtensions_UploadFile_Destination_file_name_is_required_
Looks up a localized string similar to Destination file name is required..
#### FileFolderExtensions_UploadFile_The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_
Looks up a localized string similar to The argument must be a single file name and cannot contain path characters..
#### FileFolderExtensions_UploadFile0ToFolder1
Looks up a localized string similar to Uploading file '{0}' to folder '{1}'..
#### FileFolderExtensions_UploadFileWebDav_The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_
Looks up a localized string similar to The argument must be a single file name and cannot contain path characters..
#### GraphExtensions_ErrorOccured
Looks up a localized string similar to Graph call returned the following error: {0}..
#### GraphExtensions_GroupLogoFileDoesNotExist
Looks up a localized string similar to The group logo file does not exist..
#### GraphExtensions_SendAsyncRetry
Looks up a localized string similar to Microsoft Graph API request frequency exceeded usage limits. Sleeping for {0} milliseconds before retrying..
#### GraphExtensions_SendAsyncRetryException
Looks up a localized string similar to SendAsync threw following exception: {0}..
#### ListExtensions_CreateList0Template12
Looks up a localized string similar to Creating list '{0}' from template {1}{2}..
#### ListExtensions_GetWebRelativeUrl
Looks up a localized string similar to Cannot establish web relative URL from the {0} list root folder URI and the {1} parent web URI..
#### ListExtensions_IncorrectValueFormat
Looks up a localized string similar to Value should be in the format of <id>;#<value> (Example 1: 1;#Foo Bar Example 2: 1;#Foo Bar;#2;#Bar Foo).
#### ListExtensions_SkipNoCrawlLists
Looks up a localized string similar to Skipping reindexing of the list because it's marked as a 'no crawl' list..
#### LoggingUtility_MessageWithException
Looks up a localized string similar to {0}; EXCEPTION: {{{1}}}.
#### MailUtility_SendException
Looks up a localized string similar to Mail message could not be sent. SMTP exception attempting to send. Error = {0}.
#### MailUtility_SendExceptionRethrow0
Looks up a localized string similar to Mail message could not be sent. Exception attempting to send email, rethrowing. Exception: {0}.
#### MailUtility_SendFailed
Looks up a localized string similar to Mail message could not be sent. Send completed with error {0}..
#### MailUtility_SendMailCancelled
Looks up a localized string similar to Mail message was canceled..
#### PnPMonitoredScope_Code_execution_ended
Looks up a localized string similar to Code execution scope ended.
#### PnPMonitoredScope_Code_execution_started
Looks up a localized string similar to Code execution scope started.
#### PnPMonitoredScopeExtensions_LogPropertyUpdate_Updating_property__0_
Looks up a localized string similar to Updating property {0}.
#### Provisioning_Asymmetric_Base_Templates
Looks up a localized string similar to The source site from which the template was generated had a base template ID value of {0}, while the current target site has a base template ID value of {1}. Thus, there could be potential issues while applying the template..
#### Provisioning_Connectors_Azure_FailedToInitialize
Looks up a localized string similar to Could not initialize AzureStorageConnector. Error = {0}.
#### Provisioning_Connectors_Azure_FileDeleted
Looks up a localized string similar to File {0} was deleted from Azure storage container {1}.
#### Provisioning_Connectors_Azure_FileDeleteFailed
Looks up a localized string similar to File {0} was not deleted from Azure storage container {1}. Error = {2}.
#### Provisioning_Connectors_Azure_FileDeleteNotFound
Looks up a localized string similar to File {0} was not deleted from Azure storage container {1} because it was not available.
#### Provisioning_Connectors_Azure_FileNotFound
Looks up a localized string similar to File {0} not found in Azure storage container {1}. Exception = {2}.
#### Provisioning_Connectors_Azure_FileRetrieved
Looks up a localized string similar to File {0} retrieved from Azure storage container {1}.
#### Provisioning_Connectors_Azure_FileSaved
Looks up a localized string similar to File {0} saved to Azure storage container {1}.
#### Provisioning_Connectors_Azure_FileSaveFailed
Looks up a localized string similar to File {0} was not saved to Azure storage container {1}. Error = {2}.
#### Provisioning_Connectors_FileSystem_FileDeleted
Looks up a localized string similar to File {0} deleted from folder {1}.
#### Provisioning_Connectors_FileSystem_FileDeleteFailed
Looks up a localized string similar to File {0} was not deleted from folder {1}. Error = {2}.
#### Provisioning_Connectors_FileSystem_FileDeleteNotFound
Looks up a localized string similar to File {0} was not deleted from folder {1} because it was not available.
#### Provisioning_Connectors_FileSystem_FileNotFound
Looks up a localized string similar to File {0} not found in directory {1}. Exception = {2}.
#### Provisioning_Connectors_FileSystem_FileRetrieved
Looks up a localized string similar to File {0} retrieved from folder {1}.
#### Provisioning_Connectors_FileSystem_FileSaved
Looks up a localized string similar to File {0} saved to folder {1}.
#### Provisioning_Connectors_FileSystem_FileSaveFailed
Looks up a localized string similar to File {0} was not saved to folder {1}. Error = {2}.
#### Provisioning_Connectors_OpenXML_FileDeleted
Looks up a localized string similar to File {0} deleted from folder {1}.
#### Provisioning_Connectors_OpenXML_FileDeleteFailed
Looks up a localized string similar to File {0} was not deleted from folder {1}. Error = {2}.
#### Provisioning_Connectors_OpenXML_FileDeleteNotFound
Looks up a localized string similar to File {0} was not deleted from folder {1} because it was not available.
#### Provisioning_Connectors_OpenXML_FileNotFound
Looks up a localized string similar to File {0} not found in directory {1}. Exception = {2}.
#### Provisioning_Connectors_OpenXML_FileRetrieved
Looks up a localized string similar to File {0} retrieved from folder {1}.
#### Provisioning_Connectors_OpenXML_FileSaved
Looks up a localized string similar to File {0} saved to folder {1}.
#### Provisioning_Connectors_OpenXML_FileSaveFailed
Looks up a localized string similar to File {0} was not saved to folder {1}. Error = {2}.
#### Provisioning_Connectors_SharePoint_FileDeleted
Looks up a localized string similar to File {0} deleted from site {1}, library {2}.
#### Provisioning_Connectors_SharePoint_FileDeleteFailed
Looks up a localized string similar to File {0} was not deleted from site {1}, library {2}. Error = {3}.
#### Provisioning_Connectors_SharePoint_FileDeleteNotFound
Looks up a localized string similar to File {0} was not deleted from site {1}, library {2} because it was not available.
#### Provisioning_Connectors_SharePoint_FileNotFound
Looks up a localized string similar to File {0} not found in site {1}, library {2}. Exception = {3}.
#### Provisioning_Connectors_SharePoint_FileRetrieved
Looks up a localized string similar to File {0} found in site {1}, library {2}.
#### Provisioning_Connectors_SharePoint_FileSaved
Looks up a localized string similar to File {0} saved to site {1}, library {2}.
#### Provisioning_Connectors_SharePoint_FileSaveFailed
Looks up a localized string similar to File {0} was not saved to site {1}, library {2}. Error = {3}.
#### Provisioning_Extensibility_Pipeline_BeforeInvocation
Looks up a localized string similar to Provisioning extensibility pipeline preparing to invoke, Assembly: {0}. Type {1}.
#### Provisioning_Extensibility_Pipeline_ClientCtxNull
Looks up a localized string similar to ClientContext is NULL. Unable to Invoke Extensibility Pipeline..
#### Provisioning_Extensibility_Pipeline_Exception
Looks up a localized string similar to There was an exception invoking the custom extensibility provider. Assembly: {0}, Type: {1}. Exception {2}.
#### Provisioning_Extensibility_Pipeline_Missing_AssemblyName
Looks up a localized string similar to Provider.Assembly missing value. Unable to Invoke Extensibility Pipeline..
#### Provisioning_Extensibility_Pipeline_Missing_TypeName
Looks up a localized string similar to Provider.Type missing value. Unable to Invoke Extensibility Pipeline..
#### Provisioning_Extensibility_Pipeline_Success
Looks up a localized string similar to Provisioning extensibility pipline invocation successful, Assembly {0}, Type {1}.
#### Provisioning_Extensions_ViewLocalization_Skip
Looks up a localized string similar to Skipping view localization because we're running under a user context who has a prefered language set in it's profile. This setup will not allow to add the needed localized string versions..
#### Provisioning_Extensions_WebPartLocalization_Skip
Looks up a localized string similar to Skipping web part localization because we're running under a user context who has a prefered language set in it's profile. This setup will not allow to add the needed localized string versions..
#### Provisioning_Formatter_Invalid_Template_URI
Looks up a localized string similar to The Provisioning Template URI {0} is not valid..
#### Provisioning_ObjectHandlers_Audit_SkipAuditLogTrimmingRetention
Looks up a localized string similar to Audit log trimming retention is not set because the site is configured for noscript..
#### Provisioning_ObjectHandlers_ComposedLooks_DownLoadFile_Downloading_asset___0_
Looks up a localized string similar to Downloading asset: {0}.
#### Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_ComposedLookInfoFailedToDeserialize
Looks up a localized string similar to Composed Look Information in Property Bag failed to deserialize. Falling back to detection of current composed look.
#### Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Creating_SharePointConnector
Looks up a localized string similar to Creating SharePointConnector.
#### Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Retrieving_current_composed_look
Looks up a localized string similar to Retrieving current composed look.
#### Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Using_ComposedLookInfoFromPropertyBag
Looks up a localized string similar to Using Composed Look Information from Property Bag.
#### Provisioning_ObjectHandlers_ComposedLooks_NoSiteCheck
Looks up a localized string similar to Skipping composed look handling because the site is marked as 'nosite'..
#### Provisioning_ObjectHandlers_ContentTypes_Adding_content_type_to_template___0_____1_
Looks up a localized string similar to Adding content type to template: {0} - {1}.
#### Provisioning_ObjectHandlers_ContentTypes_Adding_field__0__to_content_type
Looks up a localized string similar to Adding field {0} to content type.
#### Provisioning_ObjectHandlers_ContentTypes_Context_web_is_subweb__Skipping_content_types_
Looks up a localized string similar to Context web is subweb. Skipping content types..
#### Provisioning_ObjectHandlers_ContentTypes_Creating_new_Content_Type___0_____1_
Looks up a localized string similar to Creating new Content Type: {0} - {1}.
#### Provisioning_ObjectHandlers_ContentTypes_DocumentSet_DeltaHandling_OnHold
Looks up a localized string similar to Content Type {0} with ID {1} cannot be updated because delta handling for DocumentSets is on hold..
#### Provisioning_ObjectHandlers_ContentTypes_Field__0__exists_in_content_type
Looks up a localized string similar to Field {0} exists in content type.
#### Provisioning_ObjectHandlers_ContentTypes_InvalidDocumentSet_Update_Request
Looks up a localized string similar to Content Type {0} with ID {1} cannot be transformed into a DocumentSet.
#### Provisioning_ObjectHandlers_ContentTypes_Recreating_existing_Content_Type___0_____1_
Looks up a localized string similar to Recreating existing Content Type: {0} - {1}.
#### Provisioning_ObjectHandlers_ContentTypes_SkipCustomFormUrls
Looks up a localized string similar to Skipping custom form urls for contenttype '{0}' because we can't upload them on 'noscript' sites..
#### Provisioning_ObjectHandlers_ContentTypes_SkipDocumentSetDefaultDocuments
Looks up a localized string similar to Skipping adding default documements to document set '{0}' because this is not supported on 'noscript' sites..
#### Provisioning_ObjectHandlers_ContentTypes_Updating_existing_Content_Type___0_____1_
Looks up a localized string similar to Updating existing Content Type: {0} - {1}.
#### Provisioning_ObjectHandlers_ContentTypes_Updating_existing_Content_Type_Sealed
Looks up a localized string similar to Existing content type with Id {0} and name {1} will not be updated because it's marked as sealed and the template is not changing the sealed value..
#### Provisioning_ObjectHandlers_CustomActions_Adding_custom_action___0___to_scope_Site
Looks up a localized string similar to Adding custom action '{0}' to scope Site.
#### Provisioning_ObjectHandlers_CustomActions_Adding_custom_action___0___to_scope_Web
Looks up a localized string similar to Adding custom action '{0}' to scope Web.
#### Provisioning_ObjectHandlers_CustomActions_Adding_site_scoped_custom_action___0___to_template
Looks up a localized string similar to Adding site scoped custom action '{0}' to template.
#### Provisioning_ObjectHandlers_CustomActions_Adding_web_scoped_custom_action___0___to_template
Looks up a localized string similar to Adding web scoped custom action '{0}' to template.
#### Provisioning_ObjectHandlers_CustomActions_Removing_site_scoped_custom_action___0___from_template_because_already_available_in_base_template
Looks up a localized string similar to Removing site scoped custom action '{0}' from template because already available in base template.
#### Provisioning_ObjectHandlers_CustomActions_Removing_web_scoped_custom_action___0___from_template_because_already_available_in_base_template
Looks up a localized string similar to Removing web scoped custom action '{0}' from template because already available in base template.
#### Provisioning_ObjectHandlers_CustomActions_SkippingAddUpdateDueToNoScript
Looks up a localized string similar to Custom action '{0}' was not added/updated because the site was configured for noscript..
#### Provisioning_ObjectHandlers_ExtensibilityProviders_Calling_extensibility_callout__0_
Looks up a localized string similar to Calling extensibility callout {0}.
#### Provisioning_ObjectHandlers_ExtensibilityProviders_Calling_tokenprovider_extensibility_callout__0_
Looks up a localized string similar to Calling extensibility tokenprovider callout {0}.
#### Provisioning_ObjectHandlers_ExtensibilityProviders_callout_failed___0_____1_
Looks up a localized string similar to Extensibility callout failed: {0} : {1}.
#### Provisioning_ObjectHandlers_ExtensibilityProviders_tokenprovider_callout_failed___0_____1_
Looks up a localized string similar to Extensibility tokenprovider callout failed: {0} : {1}.
#### Provisioning_ObjectHandlers_Extraction
Looks up a localized string similar to Extraction.
#### Provisioning_ObjectHandlers_Features_Activating__0__scoped_feature__1_
Looks up a localized string similar to Activating {0} scoped feature {1}.
#### Provisioning_ObjectHandlers_Features_Deactivating__0__scoped_feature__1_
Looks up a localized string similar to Deactivating {0} scoped feature {1}.
#### Provisioning_ObjectHandlers_Fields_Adding_field__0__failed___1_____2_
Looks up a localized string similar to Adding field {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_Fields_Adding_field__0__to_site
Looks up a localized string similar to Adding field {0} to site.
#### Provisioning_ObjectHandlers_Fields_Context_web_is_subweb__skipping_site_columns
Looks up a localized string similar to Context web is subweb, skipping site columns.
#### Provisioning_ObjectHandlers_Fields_Field__0____1___exists_but_is_of_different_type__Skipping_field_
Looks up a localized string similar to Field {0} ({1}) exists but is of different type. Skipping field..
#### Provisioning_ObjectHandlers_Fields_Updating_field__0__failed___1_____2_
Looks up a localized string similar to Updating field {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_Fields_Updating_field__0__in_site
Looks up a localized string similar to Updating field {0} in site.
#### Provisioning_ObjectHandlers_Files_Adding_webpart___0___to_page
Looks up a localized string similar to Adding webpart '{0}' to page.
#### Provisioning_ObjectHandlers_Files_SkipFileUpload
Looks up a localized string similar to Skipping upload of file '{0}' to '{1}'..
#### Provisioning_ObjectHandlers_Files_Uploading_and_overwriting_existing_file__0_
Looks up a localized string similar to Uploading and overwriting existing file {0}.
#### Provisioning_ObjectHandlers_Files_Uploading_file__0_
Looks up a localized string similar to Uploading file {0}.
#### Provisioning_ObjectHandlers_FinishExtraction
Looks up a localized string similar to FINISH - Template Extraction.
#### Provisioning_ObjectHandlers_FinishProvisioning
Looks up a localized string similar to FINISH - Provisioning.
#### Provisioning_ObjectHandlers_ListInstances_Adding_list___0_____1_
Looks up a localized string similar to Adding list: {0} - {1}.
#### Provisioning_ObjectHandlers_ListInstances_Creating_field__0_
Looks up a localized string similar to Creating field {0}.
#### Provisioning_ObjectHandlers_ListInstances_Creating_field__0__failed___1_____2_
Looks up a localized string similar to Creating field {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_ListInstances_Creating_list__0_
Looks up a localized string similar to Creating list {0}.
#### Provisioning_ObjectHandlers_ListInstances_Creating_list__0__failed___1_____2_
Looks up a localized string similar to Creating list {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_ListInstances_Creating_view__0_
Looks up a localized string similar to Creating view {0}.
#### Provisioning_ObjectHandlers_ListInstances_Creating_view_failed___0_____1_
Looks up a localized string similar to Creating view failed: {0} : {1}.
#### Provisioning_ObjectHandlers_ListInstances_DraftVersionVisibility_not_applied_because_EnableModeration_is_not_set_to_true
Looks up a localized string similar to DraftVersionVisibility not applied because EnableModeration is not set to true.
#### Provisioning_ObjectHandlers_ListInstances_Field__0____1___exists_in_list__2____3___but_is_of_different_type__Skipping_field_
Looks up a localized string similar to Field {0} ({1}) exists in list {2} ({3}) but is of different type. Skipping field..
#### Provisioning_ObjectHandlers_ListInstances_Field_schema_has_no_ID_attribute___0_
Looks up a localized string similar to Field schema has no ID attribute: {0}.
#### Provisioning_ObjectHandlers_ListInstances_FieldRef_Updating_list__0_
Looks up a localized string similar to Updating list {0} with FieldRef {1}.
#### Provisioning_ObjectHandlers_ListInstances_FolderAlreadyExists
Looks up a localized string similar to Folder '{0}' already exists in parent folder '{1}'..
#### Provisioning_ObjectHandlers_ListInstances_ID_for_field_is_not_a_valid_Guid___0_
Looks up a localized string similar to ID for field is not a valid Guid: {0}.
#### Provisioning_ObjectHandlers_ListInstances_InvalidFieldReference
Looks up a localized string similar to The List {0} references site field {1} ({2}) which could not be found in the site. Use of the site field has been aborted..
#### Provisioning_ObjectHandlers_ListInstances_List__0____1____2___exists_but_is_of_a_different_type__Skipping_list_
Looks up a localized string similar to List {0} ({1}, {2}) exists but is of a different type. Skipping list..
#### Provisioning_ObjectHandlers_ListInstances_SkipAddingOrUpdatingCustomActions
Looks up a localized string similar to Skip adding/updating custom actions because the site has "noscript" enabled..
#### Provisioning_ObjectHandlers_ListInstances_Updating_field__0_
Looks up a localized string similar to Updating field {0}.
#### Provisioning_ObjectHandlers_ListInstances_Updating_field__0__failed___1_____2_
Looks up a localized string similar to Updating field {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_ListInstances_Updating_list__0_
Looks up a localized string similar to Updating list {0}.
#### Provisioning_ObjectHandlers_ListInstances_Updating_list__0__failed___1_____2_
Looks up a localized string similar to Updating list {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_ListInstancesDataRows
Looks up a localized string similar to Data Rows.
#### Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_list_item__0_
Looks up a localized string similar to Creating list item {0}.
#### Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_duplicate
Looks up a localized string similar to This row already exists and will be skipped because the IgnoreDuplicateDataRowErrors flag is set to true..
#### Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_failed___0_____1_
Looks up a localized string similar to Creating listitem failed: {0} : {1}.
#### Provisioning_ObjectHandlers_ListInstancesDataRows_Processing_data_rows_for__0_
Looks up a localized string similar to Processing data rows for {0}.
#### Provisioning_ObjectHandlers_LookupFields_LookupTargetListLookupFailed__0
Looks up a localized string similar to Unable to find lookup list with Id: {0}.
#### Provisioning_ObjectHandlers_LookupFields_Processing_lookup_fields_failed___0_____1_
Looks up a localized string similar to Processing lookup fields failed: {0} : {1}.
#### Provisioning_ObjectHandlers_Navigation_Context_web_is_not_publishing
Looks up a localized string similar to Context web does not have the publishing features enabled, skipping navigation settings.
#### Provisioning_ObjectHandlers_Navigation_missing_current_managed_navigation
Looks up a localized string similar to Missing Current Managed Navigation settings in the current template.
#### Provisioning_ObjectHandlers_Navigation_missing_current_structural_navigation
Looks up a localized string similar to Missing Current Structural Navigation settings in the current template.
#### Provisioning_ObjectHandlers_Navigation_missing_global_managed_navigation
Looks up a localized string similar to Missing Global Managed Navigation settings in the current template.
#### Provisioning_ObjectHandlers_Navigation_missing_global_structural_navigation
Looks up a localized string similar to Missing Global Structural Navigation settings in the current template.
#### Provisioning_ObjectHandlers_Navigation_SkipProvisioning
Looks up a localized string similar to Skip applying of navigation settings on NoScript sites..
#### Provisioning_ObjectHandlers_Pages_Creating_new_page__0_
Looks up a localized string similar to Creating new page {0}.
#### Provisioning_ObjectHandlers_Pages_Creating_new_page__0__failed___1_____2_
Looks up a localized string similar to Creating new page {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0_
Looks up a localized string similar to Overwriting existing page {0}.
#### Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0__failed___1_____2_
Looks up a localized string similar to Overwriting existing page {0} failed: {1} : {2}.
#### Provisioning_ObjectHandlers_Pages_SkipAddingWebParts
Looks up a localized string similar to Skipped adding web parts to page '{0}' because the site is configured for noscript..
#### Provisioning_ObjectHandlers_PersistTemplateInformation
Looks up a localized string similar to Persist Template Information.
#### Provisioning_ObjectHandlers_PropertyBagEntries_Creating_new_propertybag_entry__0__with_value__1__2_
Looks up a localized string similar to Creating new propertybag entry {0} with value {1}{2}.
#### Provisioning_ObjectHandlers_PropertyBagEntries_Overwriting_existing_propertybag_entry__0__with_value__1_
Looks up a localized string similar to Overwriting existing propertybag entry {0} with value {1}.
#### Provisioning_ObjectHandlers_Provisioning
Looks up a localized string similar to Provisioning.
#### Provisioning_ObjectHandlers_Publishing_SkipProvisioning
Looks up a localized string similar to Skip provisioning of publishing settings because the site is configured for noscript..
#### Provisioning_ObjectHandlers_RetrieveTemplateInfo
Looks up a localized string similar to Retrieve Template Info.
#### Provisioning_ObjectHandlers_SitePolicy_PolicyAdded
Looks up a localized string similar to Site policy '{0}' applied to site.
#### Provisioning_ObjectHandlers_SitePolicy_PolicyNotFound
Looks up a localized string similar to Site policy '{0}' not found.
#### Provisioning_ObjectHandlers_SiteSecurity_Add_users_failed_for_group___0_____1_____2_
Looks up a localized string similar to Add users failed for group '{0}': {1} : {2}.
#### Provisioning_ObjectHandlers_SiteSecurity_Context_web_is_subweb__skipping_site_security_provisioning
Looks up a localized string similar to Context web is subweb, skipping site security provisioning.
#### Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_
Looks up a localized string similar to Skipping label {0}, label is to set to default for language {1} while the default termstore language is also {1} or language {1} is not defined for the termstore.
#### Provisioning_ObjectHandlers_TermGroups_Wrong_Configuration
Looks up a localized string similar to The Managed Metadata Service is not properly configured. Please set a default storage location for Keywords and for column specific term sets. The TermGroups handler execution will be skipped!.
#### Provisioning_ObjectHandlers_WebSettings_SkipCustomMasterPageUpdate
Looks up a localized string similar to Skipping custom master page update because the site is configured for noscript..
#### Provisioning_ObjectHandlers_WebSettings_SkipMasterPageUpdate
Looks up a localized string similar to Skipping master page update because the site is configured for noscript..
#### Provisioning_ObjectHandlers_WebSettings_SkipNoCrawlUpdate
Looks up a localized string similar to Skipping NoCrawl update because the site is configured for noscript..
#### Provisioning_Providers_XML_InvalidFileFormat
Looks up a localized string similar to Cannot process XML file {0}..
#### ProvisioningExtensions_ErrorProvisioningModule0File1
Looks up a localized string similar to Error provisioning module '{0}' file '{1}'. Error = {2}.
#### ProvisioningExtensions_ProvisionElementFile_Path_to_the_element_file_is_required
Looks up a localized string similar to Path to the element file is required.
#### ProvisioningExtensions_ProvisionElementFile0
Looks up a localized string similar to Provisioning Elements file '{0}'..
#### ProvisioningExtensions_ProvisionElementXml_Expected_element__Elements__
Looks up a localized string similar to Expected element 'Elements'..
#### ProvisioningExtensions_ProvisionFileInternal_Expected_element__File__
Looks up a localized string similar to Expected element 'File'..
#### ProvisioningExtensions_ProvisionModuleInternal_Expected_element__Module__
Looks up a localized string similar to Expected element 'Module'..
#### SecurityExtensions_Error_VisitingSecurableObject
Looks up a localized string similar to Something wrong happened while visiting securable object: {0}, details: {1}.
#### SecurityExtensions_Info_VisitingSecurableObject
Looks up a localized string similar to Visiting securable object: {0}.
#### SecurityExtensions_Warning_SkipFurtherVisitingForTooManyChildObjects
Looks up a localized string similar to Skip visiting the child securable objects for {0}, unique_permission_item_count = {1}, leaf_breadth_limit = {2}.
#### Service_RegistrationFailed
Looks up a localized string similar to Service registration for {0} using endpoint {1} and cachekey {2} failed..
#### Services_AccessDenied
Looks up a localized string similar to Service requestor is not registered: access denied.
#### Services_CookieWithCachKeyNotFound
Looks up a localized string similar to The cookie with the cachekey was not found...nothing can be retrieved from cache, so no clientcontext can be created..
#### Services_Registered
Looks up a localized string similar to Service {0} has been registered for endpoint {1} using cachekey {2}..
#### Services_TokenRefreshed
Looks up a localized string similar to Token for cachekey {0} and hostweburl {1} has been refreshed..
#### SiteToTemplateConversion_ApplyRemoteTemplate_OverwriteSystemPropertyBagValues_is_to_true
Looks up a localized string similar to OverwriteSystemPropertyBagValues is to true.
#### SiteToTemplateConversion_Base_template_available___0_
Looks up a localized string similar to Base template available: {0}.
#### SiteToTemplateConversion_IncludeAllTermGroups_is_set_to_true
Looks up a localized string similar to IncludeAllTermGroups is set to true.
#### SiteToTemplateConversion_IncludeSiteCollectionTermGroup_is_set_to_true
Looks up a localized string similar to IncludeSiteCollectionTermGroup is set to true.
#### SiteToTemplateConversion_MessagesDelegate_registered
Looks up a localized string similar to MessagesDelegate registered.
#### SiteToTemplateConversion_PersistBrandingFiles_is_set_to_true
Looks up a localized string similar to PersistBrandingFiles is set to true.
#### SiteToTemplateConversion_PersistComposedLookFiles_is_set_to_true
Looks up a localized string similar to PersistComposedLookFiles is set to true.
#### SiteToTemplateConversion_ProgressDelegate_registered
Looks up a localized string similar to ProgressDelegate registered.
#### TaxonomyExtension_CreateTerm01UnderParent2
Looks up a localized string similar to Creating term '{0}|{1}' under parent '{2}'..
#### TaxonomyExtension_CreateTermGroup0InStore1
Looks up a localized string similar to Creating term group '{0}' in term store '{1}'..
#### TaxonomyExtension_CreateTermSet0InGroup1
Looks up a localized string similar to Creating term set '{0}' in term group '{1}'..
#### TaxonomyExtension_DeleteTerm01
Looks up a localized string similar to Deleting term '{0}|{1}'..
#### TaxonomyExtension_ExceptionUpdateDescriptionGroup01
Looks up a localized string similar to Error setting description for term group '{0}' ({1}). Error = {2}.
#### TaxonomyExtension_ExceptionUpdateDescriptionSet01
Looks up a localized string similar to Error setting description for term set '{0}' ({1}). Error = {2}.
#### TaxonomyExtension_ImportErrorDeleteId0Line1
Looks up a localized string similar to Error encountered during import when attempting to delete invalid term with id {0} on line {1}. Error = {2}.
#### TaxonomyExtension_ImportErrorDescription0Line1
Looks up a localized string similar to Error encountered during import. The description '{0}' on line {1} is not valid..
#### TaxonomyExtension_ImportErrorName0Line1
Looks up a localized string similar to Error encountered during import. The name '{0}' is not valid on line {1}..
#### TaxonomyExtension_ImportErrorTaggingLine0
Looks up a localized string similar to Error encountered during import. The available for tagging entry on line {0} is not valid..
#### TaxonomyExtension_ImportTermSet
Looks up a localized string similar to Importing term set from file stream..
#### TaxonomyExtension_TermGroup0Id1DoesNotMatchSpecifiedId2
Looks up a localized string similar to Term group '{0}' ID ({1}) does not match specified ID ({2})..
#### TaxonomyExtension_TermSet0Id1DoesNotMatchSpecifiedId2
Looks up a localized string similar to Term set '{0}' ID ({1}) does not match specified ID ({2})..
#### TaxonomyExtensions_Field_Is_Not_Multivalues
Looks up a localized string similar to The taxonomy field {0} does not support multiple values..
#### TaxonomyExtensions_ImportTermSet_File_path_is_required_
Looks up a localized string similar to File path is required..
#### TaxonomyExtensions_ImportTermSetImplementation_Invalid_CSV_format__was_expecting_a_comma_in_the_first__header__line_
Looks up a localized string similar to Invalid CSV format; was expecting a comma in the first (header) line..
#### TenantExtensions_ClosedContextWarning
Looks up a localized string similar to ClientContext gets closed after action is completed. Calling ExecuteQuery again returns an error. Verify that you have an open ClientContext object. Error = {0}.
#### TenantExtensions_SetLockState
Looks up a localized string similar to SetSiteLockState: Current: {0} Target: {1}.
#### TenantExtensions_UnknownExceptionAccessingSite
Looks up a localized string similar to Could not determine if site exists in tenant. Error = {0}.
#### TimerJob_AddSite_Done
Looks up a localized string similar to Site {0} url/wildcard added.
#### TimerJob_AddSite_InvalidUrl
Looks up a localized string similar to Site url ({0}) contains invalid characters.
#### TimerJob_Authentication_AppOnly
Looks up a localized string similar to Timer job authentication set to type App-Only with clientId {0}.
#### TimerJob_Authentication_AzureADAppOnly
Looks up a localized string similar to Timer job authentication set to type Azure AD App-Only with clientId {0} and certificate {1}.
#### TimerJob_Authentication_Network
Looks up a localized string similar to Timer job authentication set to type NetworkCredentials with user {0} in domain {1}.
#### TimerJob_Authentication_O365
Looks up a localized string similar to Timer job authentication set to type Office 365 with user {0}.
#### TimerJob_Authentication_RetrieveFromCredMan
Looks up a localized string similar to Retrieving credetials with name {0} from the Windows Credential Manager.
#### TimerJob_Authentication_RetrieveFromCredManFailed
Looks up a localized string similar to Failed to retrieve credential manager credentials with name {0} or retrieved credentials don't have user or password set.
#### TimerJob_Authentication_TenantAdmin
Looks up a localized string similar to Tenant admin site set to {0}..
#### TimerJob_ClearAddedSites
Looks up a localized string similar to All added sites are cleared.
#### TimerJob_Clone
Looks up a localized string similar to Timer job {0} settings cloned to timer job {0}.
#### TimerJob_Constructor
Looks up a localized string similar to Timer job constructed with name {0}, version {1}.
#### TimerJob_DoWork_Done
Looks up a localized string similar to Work for site {0} done.
#### TimerJob_DoWork_NoEventHandler
Looks up a localized string similar to No event receiver connected to the TimerJobRun event.
#### TimerJob_DoWork_Start
Looks up a localized string similar to Doing work for site {0}.
#### TimerJob_Enumeration_Network
Looks up a localized string similar to Enumeration credentials specified for on-premises enumeration with user {0} and demain {1}.
#### TimerJob_Enumeration_NoDomain
Looks up a localized string similar to No domain specified that can be used for site enumeration. Use the SetEnumerationNetworkCredentials method to provide credentials as app-only does not work with search.
#### TimerJob_Enumeration_NoPassword
Looks up a localized string similar to No password specified that can be used for site enumeration. Use the SetEnumeration... method to provide credentials as app-only does not work with search.
#### TimerJob_Enumeration_NoUser
Looks up a localized string similar to No user specified that can be used for site enumeration. Use the SetEnumeration... method to provide credentials as app-only does not work with search.
#### TimerJob_Enumeration_O365
Looks up a localized string similar to Enumeration credentials specified for Office 365 enumeration with user {0}.
#### TimerJob_ExpandSite_EatException
Looks up a localized string similar to Eating exception {0} for site {1}.
#### TimerJob_ExpandSubSites
Looks up a localized string similar to ExpandSubSites set to {0}.
#### TimerJob_ManageState
Looks up a localized string similar to Manage state set to {0}.
#### TimerJob_MaxThread1
Looks up a localized string similar to If you only want 1 thread then set the UseThreading property to false.
#### TimerJob_MaxThread100
Looks up a localized string similar to You cannot use more than 100 threads.
#### TimerJob_MaxThreadLessThan1
Looks up a localized string similar to Number of threads must be between 2 and 100.
#### TimerJob_MaxThreadSet
Looks up a localized string similar to MaximumThreads set to {0}.
#### TimerJob_OnTimerJobRun_CallEventHandler
Looks up a localized string similar to Calling the eventhandler for site {0}.
#### TimerJob_OnTimerJobRun_CallEventHandlerDone
Looks up a localized string similar to Eventhandler called for site {0}.
#### TimerJob_OnTimerJobRun_Error
Looks up a localized string similar to Error during timerjob execution of site {0}. Exception message = {1}.
#### TimerJob_OnTimerJobRun_PrevRunRead
Looks up a localized string similar to Timerjob for site {1}, PreviousRun = {0}.
#### TimerJob_OnTimerJobRun_PrevRunSet
Looks up a localized string similar to Set Timerjob for site {1}, PreviousRun to {0}.
#### TimerJob_OnTimerJobRun_PrevRunSuccessRead
Looks up a localized string similar to Timerjob for site {1}, PreviousRunSuccessful = {0}.
#### TimerJob_OnTimerJobRun_PrevRunSuccessSet
Looks up a localized string similar to Set Timerjob for site {1}, PreviousRunSuccessful to {0}.
#### TimerJob_OnTimerJobRun_PrevRunVersionRead
Looks up a localized string similar to Timerjob for site {1}, PreviousRunVersion = {0}.
#### TimerJob_OnTimerJobRun_PrevRunVersionSet
Looks up a localized string similar to Set Timerjob for site {1}, PreviousRunVersion to {0}.
#### TimerJob_OnTimerJobRun_PropertiesRead
Looks up a localized string similar to Timerjob properties read using key {0} for site {1}.
#### TimerJob_OnTimerJobRun_PropertiesSet
Looks up a localized string similar to Timerjob properties written using key {0} for site {1}.
#### TimerJob_Realm
Looks up a localized string similar to Realm set to {0}.
#### TimerJob_ResolveSites_Done
Looks up a localized string similar to Resolving sites done, sub sites have been expanded.
#### TimerJob_ResolveSites_DoneNoExpansionNeeded
Looks up a localized string similar to Resolving sites done, no expansion needed.
#### TimerJob_ResolveSites_LaunchThreadPerBatch
Looks up a localized string similar to Expand subsites by launching a thread for each of the {0} work batches.
#### TimerJob_ResolveSites_ResolveSite
Looks up a localized string similar to Resolving wildcard site {0}.
#### TimerJob_ResolveSites_ResolveSiteDone
Looks up a localized string similar to Done resolving wildcard site {0}.
#### TimerJob_ResolveSites_SequentialExpandDone
Looks up a localized string similar to Done sequentially expanding all sites.
#### TimerJob_ResolveSites_Started
Looks up a localized string similar to Resolving sites started.
#### TimerJob_ResolveSites_StartSequentialExpand
Looks up a localized string similar to Start sequentially expanding all sites.
#### TimerJob_ResolveSites_ThreadLaunched
Looks up a localized string similar to Thread started to expand a batch of {0} sites.
#### TimerJob_ResolveSites_ThreadsAreDone
Looks up a localized string similar to Done waiting for all site expanding threads.
#### TimerJob_Run_AfterResolveAddedSites
Looks up a localized string similar to After calling the virtual ResolveAddedSites method. Current count of site url's = {0}.
#### TimerJob_Run_AfterUpdateAddedSites
Looks up a localized string similar to After calling the virtual UpdateAddedSites method. Current count of site url's = {0}.
#### TimerJob_Run_BeforeResolveAddedSites
Looks up a localized string similar to Before calling the virtual ResolveAddedSites method. Current count of site url's = {0}.
#### TimerJob_Run_BeforeStartWorkBatches
Looks up a localized string similar to Ready to start a thread for each of the {0} work batches.
#### TimerJob_Run_BeforeUpdateAddedSites
Looks up a localized string similar to Before calling the virtual UpdateAddedSites method. Current count of site url's = {0}.
#### TimerJob_Run_Done
Looks up a localized string similar to Run of timer job has ended.
#### TimerJob_Run_DoneProcessingWorkBatches
Looks up a localized string similar to Done processing the {0} work batches.
#### TimerJob_Run_NoSites
Looks up a localized string similar to Job does not have sites to process, bailing out.
#### TimerJob_Run_ProcessSequentially
Looks up a localized string similar to Ready to process each of the {0} sites in a sequential manner.
#### TimerJob_Run_ProcessSequentiallyDone
Looks up a localized string similar to Done with sequentially processing each of the {0} sites.
#### TimerJob_Run_Started
Looks up a localized string similar to Run of timer job has started.
#### TimerJob_Run_ThreadLaunched
Looks up a localized string similar to Thread launched for processing {0} sites.
#### TimerJob_SharePointVersion
Looks up a localized string similar to SharePointVersion set to {0}.
#### TimerJob_SharePointVersion_Versions
Looks up a localized string similar to SharePoint version must be 15 or 16.
#### TimerJob_UseThreading
Looks up a localized string similar to UseThreading set to {0}.
#### WebExtensions_CreatePublishingImageRendition
Looks up a localized string similar to Creating Image Rendition '{0}' of width '{1}' and height '{2}'..
#### WebExtensions_CreatePublishingImageRendition_Error
Looks up a localized string similar to Unable to create Image Rendition '{0}'..
#### WebExtensions_CreateWeb
Looks up a localized string similar to Creating web '{0}' with template '{1}'..
#### WebExtensions_DeleteWeb
Looks up a localized string similar to Deleting web '{0}'..
#### WebExtensions_InstallSolution
Looks up a localized string similar to Installing sandbox solution '{0}' to '{1}'..
#### WebExtensions_RemoveAppInstance
Looks up a localized string similar to Removing app '{0}' instance {1}..
#### WebExtensions_RequestAccessEmailLimitExceeded
Looks up a localized string similar to Request access email addresses exceed 255 characters. Skipping: {0}.
#### WebExtensions_SiteSearchUnhandledException
Looks up a localized string similar to Site search error. Error = {0}.
#### WebExtensions_UninstallSolution
Looks up a localized string similar to Removing sandbox solution '{0}'..

## Core.Framework.Provisioning.Providers.Xml.Resolvers.TermSetFromModelToSchemaTypeResolver
            
Resolves a TermSet collection type from Domain Model to Schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.ExpressionCollectionValueResolver
            
Resolve collection from model to schema with expression
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.ExpressionCollectionValueResolver`1
            
Resolve collection from schema to model with expression
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromAuditFlagsToArrayResolver
            
Resolves an enum bit mask of AuditFlags into an array of Strings
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromArrayToAuditFlagsResolver
            
Resolves an array of Strings into an enum bit mask of AuditFlags
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromBasePermissionsToStringValueResolver
            
Resolves a Decimal value into a Double
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromNullableToSpecifiedValueResolver`1
            
Resolves a Decimal value into a Double
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.PageLayoutsFromModelToSchemaTypeResolver
            
Resolves a list of Views from Schema to Domain Model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.PageLayoutsFromSchemaToModelTypeResolver
            
Resolves a list of Views from Schema to Domain Model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.NavigationFromModelToSchemaTypeResolver
            
Resolves a Navigation type from model to schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.NavigationFromSchemaToModelTypeResolver
            
Resolves a Navigation type from schema to model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.NavigationNodeFromModelToSchemaTypeResolver
            
Type resolver for Navigation Node from model to schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.NavigationNodeFromSchemaToModelTypeResolver
            
Type resolver for Navigation Node from schema to model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.RoleAssigmentsFromModelToSchemaTypeResolver
            
Resolves a collection type from Domain Model to Schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.RoleAssigmentsFromSchemaToModelTypeResolver
            
Resolves a collection type from Domain Model to Schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.PropertyObjectTypeResolver`1
            
Typed vesion of PropertyObjectTypeResolver
            
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.PropertyObjectTypeResolver
            
Resolves a collection type from Domain Model to Schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromStringToBasePermissionsValueResolver
            
Resolves a Decimal value into a Double
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromStringToEnumValueResolver
            
Resolves a Decimal value into a Double
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.DocumentSetTemplateFromModelToSchemaTypeResolver
            
Resolves a Template Parameter type from Domain Model to Schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.SecurityFromModelToSchemaTypeResolver
            
Resolver for Security settings from model to schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.XmlAnyFromModeToSchemalValueResolver
            
Resolves a Dictionary into an Array of objects
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.CollectionFromModelToSchemaTypeResolver
            
Resolves a collection type from Domain Model to Schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.RemoveExistingViewsFromSchemaToModelValueResolver
            
Resolves a list of Views from Schema to Domain Model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.ListViewsFromSchemaToModelTypeResolver
            
Resolves a list of Views from Schema to Domain Model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.CollectionFromSchemaToModelTypeResolver
            
Resolves a type from Schema to Domain Model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromArrayToDictionaryValueResolver`2
            
Resolves an Array of object into a Dictionary
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromStringToGuidValueResolver
            
Resolves a Decimal value into a Double
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromDecimalToDoubleValueResolver
            
Resolves a Decimal value into a Double
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromDictionaryToArrayValueResolver`2
            
Resolves a Dictionary into an Array of objects
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.FromDoubleToDecimalValueResolver
            
Resolves a Double value into a Decimal
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.ResolversExtensions
            
Extension type for resolvers
        
### Methods


#### GetPublicInstancePropertyValue(System.Object,System.String)
Retrieves the value of a public, instance property
> ##### Parameters
> **source:** The source object

> **propertyName:** The property name, case insensitive

> ##### Return value
> The property value, if any

#### GetPublicInstanceProperty(System.Object,System.String)
Retrieves a public, instance property
> ##### Parameters
> **source:** The source object

> **propertyName:** The property name, case insensitive

> ##### Return value
> The property, if any

#### SetPublicInstancePropertyValue(System.Object,System.String,System.Object)
Sets the value of a public, instance property
> ##### Parameters
> **source:** The source object

> **propertyName:** The property name, case insensitive


## Core.Framework.Provisioning.Providers.Xml.Resolvers.SecurityFromSchemaToModelTypeResolver
            
Resolver for Security settings from schema to model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.TemplateParameterFromModelToSchemaTypeResolver
            
Resolves a Template Parameter type from Domain Model to Schema
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.TemplateParameterFromSchemaToModelTypeResolver
            
Resolves a Template Parameter type from Schema to Domain Model
        

## Core.Framework.Provisioning.Providers.Xml.Resolvers.XmlAnyFromSchemaToModelValueResolver
            
Resolves a Dictionary into an Array of objects
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.ExtensibilityHandlersSerializer
            
Class to serialize/deserialize the Providers for Extensibility
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.AddInsSerializer
            
Class to serialize/deserialize the AddIns
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.PublishingSerializer
            
Class to serialize/deserialize the Publishing settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.ComposedLookSerializer
            
Class to serialize/deserialize the ComposedLook settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.NavigationSerializer
            
Class to serialize/deserialize the Navigation settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.SecuritySerializer
            
Class to serialize/deserialize the Security settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.FeaturesSerializer
            
Class to serialize/deserialize the Features
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.DirectoriesSerializer
            
Class to serialize/deserialize the Directories
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.WorkflowsActionsSerializer
            
Class to serialize/deserialize the Workflows
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.AuditSettingsSerializer
            
Class to serialize/deserialize the Audit Settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.SiteColumnsSerializer
            
Class to serialize/deserialize the Site Columns
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.SupportedUILanguagesSerializer
            
Class to serialize/deserialize the Supported UI Languages
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.TermGroupsSerializer
            
Class to serialize/deserialize the Term Groups
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.PagesSerializer
            
Class to serialize/deserialize the Pages
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.FilesSerializer
            
Class to serialize/deserialize the Files
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.CustomActionsSerializer
            
Class to serialize/deserialize the Custom Actions
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.TemplateBasePropertiesSerializer
            
Class to serialize/deserialize the Base Properties of a Template
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.RegionalSettingsSerializer
            
Class to serialize/deserialize the Regional Settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.WebSettingsSerializer
            
Class to serialize/deserialize the Web Settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.ContentTypesSerializer
            
Class to serialize/deserialize the Content Types
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.ListInstancesSerializer
            
Class to serialize/deserialize the List Instances
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.LocalizationsSerializer
            
Class to serialize/deserialize the Localization Settings
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.PropertyBagPropertiesSerializer
            
Class to serialize/deserialize the Property Bag Properties
        

## Core.Framework.Provisioning.Providers.Xml.Serializers.TemplateParametersSerializer
            
Class to serialize/deserialize the Parameters of the Template
        

## Core.Framework.Provisioning.Providers.Xml.IPnPSchemaSerializer
            
Basic interface for every Schema Serializer type
        
### Properties

#### Name
Provides the name of the serializer type
### Methods


#### Deserialize(System.Object,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate)
The method to deserialize an XML Schema based object into a Domain Model object
> ##### Parameters
> **persistence:** The persistence layer object

> **template:** The PnP Provisioning Template object


#### Serialize(OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,System.Object)
The method to serialize a Domain Model object into an XML Schema based object
> ##### Parameters
> **template:** The PnP Provisioning Template object

> **persistence:** The persistence layer object


## Core.Framework.Provisioning.Providers.Xml.IResolver
            
Basic interface for all the resolver types
        
### Properties

#### Name
Provides the name of the Resolver

## Core.Framework.Provisioning.Providers.Xml.ITypeResolver
            
Handles custom type resolving rules for PnPObjectsMapper
        
### Methods


#### Resolve(System.Object,System.Collections.Generic.Dictionary{System.String,OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.IResolver},System.Boolean)
Resolves a source type into a result
> ##### Parameters
> **source:** The full source object to resolve

> **resolvers:** 

> **recursive:** Defines whether to apply the mapping recursively, optional and by default false


## Core.Framework.Provisioning.Providers.Xml.IValueResolver
            
Handles custom value resolving rules for PnPObjectsMapper
        
### Methods


#### Resolve(System.Object,System.Object,System.Object)
Resolves a source value into a result
> ##### Parameters
> **source:** The full source object to resolve

> **destination:** The full destination object to resolve

> **sourceValue:** The source value to resolve

> ##### Return value
> The resolved value

## Core.Framework.Provisioning.Providers.Xml.IXMLSchemaFormatter
            
Interface for template formatters that read and write XML documents
        
### Properties

#### NamespaceUri
The URI of the target XML Namespace
#### NamespacePrefix
The default namespace prefix of the target XML Namespace

## Core.Framework.Provisioning.Providers.Xml.PnPBaseSchemaSerializer`1
            
Base class for every Schema Serializer
        
### Methods


#### CreateSelectorLambda(System.Type,System.String)
Protected method to create a Lambda Expression like: i => i.Property
> ##### Parameters
> **targetType:** The Type of the .NET property to apply the Lambda Expression to

> **propertyName:** The name of the property of the target object

> ##### Return value
> 

## Core.Framework.Provisioning.Providers.Xml.PnPObjectsMapper
            
Utility class that maps one object to another
        
### Methods


#### MapProperties``1(``0,System.Object,System.Collections.Generic.Dictionary{System.Linq.Expressions.Expression{System.Func{``0,System.Object}},OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.IResolver},System.Boolean)
Maps the properties of a typed source object, to the properties of an untyped destination object
> ##### Parameters
> **source:** The source object

> **destination:** The destination object

> **resolverExpressions:** Any custom resolver, optional

> **recursive:** Defines whether to apply the mapping recursively, optional and by default false


#### MapProperties``1(System.Object,``0,System.Collections.Generic.Dictionary{System.Linq.Expressions.Expression{System.Func{``0,System.Object}},OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.IResolver},System.Boolean)
Maps the properties of an untyped source object object, to the properties of a typed destination object
> ##### Parameters
> **source:** The source object

> **destination:** The destination object

> **resolverExpressions:** Any custom resolver, optional

> **recursive:** Defines whether to apply the mapping recursively, optional and by default false


#### MapProperties(System.Object,System.Object,System.Collections.Generic.Dictionary{System.String,OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.IResolver},System.Boolean)
Maps the properties of a source object, to the properties of a destination object
> ##### Parameters
> **source:** The source object

> **destination:** The destination object

> **resolvers:** Any custom resolver, optional

> **recursive:** Defines whether to apply the mapping recursively, optional and by default false


#### MapObjects``1(System.Object,OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ITypeResolver,System.Collections.Generic.Dictionary{System.Linq.Expressions.Expression{System.Func{``0,System.Object}},OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.IResolver},System.Boolean)
Maps a source object, into a destination object
> ##### Parameters
> **source:** The source object

> **resolver:** A custom resolver

> **resolverExpressions:** Any custom resolver, optional

> **recursive:** Defines whether to apply the mapping recursively, optional and by default false

> ##### Return value
> The mapped destination object

#### MapObjects(System.Object,OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ITypeResolver,System.Collections.Generic.Dictionary{System.String,OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.IResolver},System.Boolean)
Maps a source object, into a destination object
> ##### Parameters
> **source:** The source object

> **resolver:** A custom resolver

> **resolvers:** Any custom resolver, optional

> **recursive:** Defines whether to apply the mapping recursively, optional and by default false

> ##### Return value
> The mapped destination object

#### ConvertExpressionsToResolvers``1(System.Collections.Generic.Dictionary{System.Linq.Expressions.Expression{System.Func{``0,System.Object}},OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.IResolver})
Transforms a Dictionary of IValueResolver instances by Expression into a Dictionary by String (property name)
> ##### Parameters
> **resolverExpressions:** The Dictionary to transform

> ##### Return value
> The transformed dictionary

## Core.Framework.Provisioning.Providers.Xml.PnPSerializationScope
            
Internal class to handle a Provisioning Template serialization scope
        
### Methods


#### Dispose
Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.

## Core.Framework.Provisioning.Providers.Xml.V201503.SharePointProvisioningTemplate
            

        
### Properties

#### SitePolicy

#### PropertyBagEntries

#### Security

#### SiteFields

#### ContentTypes

#### Lists

#### Features

#### CustomActions

#### Files

#### ComposedLook

#### Providers

#### ID

#### Version

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201503.PropertyBagEntry
            

        
### Properties

#### Key

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201503.Provider
            

        
### Properties

#### Configuration

#### Enabled

#### Assembly

#### Type


## Core.Framework.Provisioning.Providers.Xml.V201503.ComposedLook
            

        
### Properties

#### Name

#### ColorFile

#### FontFile

#### BackgroundFile

#### MasterPage

#### SiteLogo

#### AlternateCSS

#### Version

#### VersionSpecified


## Core.Framework.Provisioning.Providers.Xml.V201503.File
            

        
### Properties

#### Src

#### Folder

#### Overwrite


## Core.Framework.Provisioning.Providers.Xml.V201503.CustomAction
            

        
### Properties

#### Name

#### Description

#### Group

#### Location

#### Title

#### Sequence

#### SequenceSpecified

#### Rights

#### RightsSpecified

#### Url

#### Enabled

#### ScriptBlock

#### ImageUrl

#### ScriptSrc


## Core.Framework.Provisioning.Providers.Xml.V201503.Feature
            

        
### Properties

#### ID

#### Deactivate

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201503.FieldRef
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201503.ContentTypeBinding
            

        
### Properties

#### ContentTypeID

#### Default


## Core.Framework.Provisioning.Providers.Xml.V201503.ListInstance
            

        
### Properties

#### ContentTypeBindings

#### Views

#### Fields

#### FieldRefs

#### Title

#### Description

#### DocumentTemplate

#### OnQuickLaunch

#### TemplateType

#### Url

#### EnableVersioning

#### MinorVersionLimit

#### MinorVersionLimitSpecified

#### MaxVersionLimit

#### MaxVersionLimitSpecified

#### RemoveDefaultContentType

#### ContentTypesEnabled

#### Hidden

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201503.ListInstanceViews
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201503.ListInstanceFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201503.User
            

        
### Properties

#### Name


## Core.Framework.Provisioning.Providers.Xml.V201503.SharePointProvisioningTemplateSecurity
            

        
### Properties

#### AdditionalAdministrators

#### AdditionalOwners

#### AdditionalMembers

#### AdditionalVisitors


## Core.Framework.Provisioning.Providers.Xml.V201503.SharePointProvisioningTemplateSiteFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201503.SharePointProvisioningTemplateContentTypes
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201503.SharePointProvisioningTemplateFeatures
            

        
### Properties

#### SiteFeatures

#### WebFeatures


## Core.Framework.Provisioning.Providers.Xml.V201503.SharePointProvisioningTemplateCustomActions
            

        
### Properties

#### SiteCustomActions

#### WebCustomActions


## Core.Framework.Provisioning.Providers.Xml.V201505.Provisioning
            

        
### Properties

#### Preferences

#### Templates

#### Sequence

#### ImportSequence

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.Preferences
            

        
### Properties

#### Parameters

#### Version

#### Author

#### Generator

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.PreferencesParameter
            

        
### Properties

#### Key

#### Required

#### Text


## Core.Framework.Provisioning.Providers.Xml.V201505.ImportSequence
            

        
### Properties

#### File


## Core.Framework.Provisioning.Providers.Xml.V201505.TermStore
            

        
### Fields

#### 

#### 

### Properties

#### TermGroup

#### Scope


## Core.Framework.Provisioning.Providers.Xml.V201505.TermGroup
            

        
### Properties

#### TermSets

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201505.TermSet
            

        
### Properties

#### CustomProperties

#### Terms

#### Language

#### LanguageSpecified

#### IsOpenForTermCreation

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.StringDictionaryItem
            

        
### Properties

#### Key

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201505.PropertyBagEntry
            

        
### Properties

#### Indexed


## Core.Framework.Provisioning.Providers.Xml.V201505.Term
            

        
### Fields

#### 

#### 

### Properties

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### Terms

#### Labels

#### CustomProperties

#### LocalCustomProperties

#### Language

#### LanguageSpecified

#### CustomSortOrder

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.TermTerms
            

        
### Properties

#### Items


## Core.Framework.Provisioning.Providers.Xml.V201505.TermLabelsLabel
            

        
### Properties

#### Language

#### Value

#### IsDefaultForLanguage


## Core.Framework.Provisioning.Providers.Xml.V201505.TermSetItem
            

        
### Properties

#### Owner

#### Description

#### IsAvailableForTagging


## Core.Framework.Provisioning.Providers.Xml.V201505.TaxonomyItem
            

        
### Properties

#### Name

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201505.TermStoreScope
            

        
### Fields

#### Default

#### Current


## Core.Framework.Provisioning.Providers.Xml.V201505.Site
            

        
### Properties

#### Templates

#### Title

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### UseSamePermissionsAsParentSite

#### UseSamePermissionsAsParentSiteSpecified

#### Url

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.Templates
            

        
### Properties

#### ProvisioningTemplateFile

#### ProvisioningTemplateReference

#### ProvisioningTemplate

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201505.ProvisioningTemplateFile
            

        
### Properties

#### File

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201505.ProvisioningTemplateReference
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201505.ProvisioningTemplate
            

        
### Properties

#### 

#### 

#### 

#### SitePolicy

#### PropertyBagEntries

#### Security

#### SiteFields

#### ContentTypes

#### Lists

#### Features

#### CustomActions

#### Files

#### Pages

#### TermGroups

#### ComposedLook

#### Providers

#### ID

#### Version

#### VersionSpecified

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.ProvisioningTemplateSecurity
            

        
### Properties

#### AdditionalAdministrators

#### AdditionalOwners

#### AdditionalMembers

#### AdditionalVisitors


## Core.Framework.Provisioning.Providers.Xml.V201505.User
            

        
### Properties

#### Name


## Core.Framework.Provisioning.Providers.Xml.V201505.ProvisioningTemplateSiteFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201505.ContentType
            

        
### Properties

#### FieldRefs

#### DocumentTemplate

#### ID

#### Name

#### Description

#### Group

#### Hidden

#### Sealed

#### ReadOnly

#### Overwrite

#### AnyAttr

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.ContentTypeFieldRef
            

        

## Core.Framework.Provisioning.Providers.Xml.V201505.FieldRefBase
            

        
### Properties

#### ID

#### Name

#### Required

#### Hidden


## Core.Framework.Provisioning.Providers.Xml.V201505.ListInstanceFieldRef
            

        
### Properties

#### DisplayName


## Core.Framework.Provisioning.Providers.Xml.V201505.ContentTypeDocumentTemplate
            

        
### Properties

#### TargetName


## Core.Framework.Provisioning.Providers.Xml.V201505.ListInstance
            

        
### Properties

#### 

#### ContentTypeBindings

#### Views

#### Fields

#### FieldRefs

#### DataRows

#### Title

#### Description

#### DocumentTemplate

#### OnQuickLaunch

#### TemplateType

#### Url

#### EnableVersioning

#### EnableMinorVersions

#### EnableModeration

#### MinorVersionLimit

#### MinorVersionLimitSpecified

#### MaxVersionLimit

#### MaxVersionLimitSpecified

#### DraftVersionVisibility

#### DraftVersionVisibilitySpecified

#### RemoveExistingContentTypes

#### TemplateFeatureID

#### ContentTypesEnabled

#### Hidden

#### EnableAttachments

#### EnableFolderCreation

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.ContentTypeBinding
            

        
### Properties

#### ContentTypeID

#### Default


## Core.Framework.Provisioning.Providers.Xml.V201505.ListInstanceViews
            

        
### Properties

#### Any

#### RemoveExistingViews


## Core.Framework.Provisioning.Providers.Xml.V201505.ListInstanceFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201505.DataValue
            

        
### Properties

#### FieldName

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201505.ProvisioningTemplateFeatures
            

        
### Properties

#### SiteFeatures

#### WebFeatures


## Core.Framework.Provisioning.Providers.Xml.V201505.Feature
            

        
### Properties

#### ID

#### Deactivate

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201505.ProvisioningTemplateCustomActions
            

        
### Properties

#### SiteCustomActions

#### WebCustomActions


## Core.Framework.Provisioning.Providers.Xml.V201505.CustomAction
            

        
### Properties

#### CommandUIExtension

#### Name

#### Description

#### Group

#### Location

#### Title

#### Sequence

#### SequenceSpecified

#### Rights

#### RightsSpecified

#### Url

#### Enabled

#### ScriptBlock

#### ImageUrl

#### ScriptSrc

#### 


## Core.Framework.Provisioning.Providers.Xml.V201505.CustomActionCommandUIExtension
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201505.File
            

        
### Properties

#### Properties

#### WebParts

#### Src

#### Folder

#### Overwrite


## Core.Framework.Provisioning.Providers.Xml.V201505.WebPartPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Zone

#### Order


## Core.Framework.Provisioning.Providers.Xml.V201505.Page
            

        
### Properties

#### WebParts

#### Url

#### Overwrite

#### Layout

#### WelcomePage


## Core.Framework.Provisioning.Providers.Xml.V201505.WikiPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Row

#### Column


## Core.Framework.Provisioning.Providers.Xml.V201505.WikiPageLayout
            

        
### Fields

#### OneColumn

#### OneColumnSidebar

#### TwoColumns

#### TwoColumnsHeader

#### TwoColumnsHeaderFooter

#### ThreeColumns

#### ThreeColumnsHeader

#### ThreeColumnsHeaderFooter


## Core.Framework.Provisioning.Providers.Xml.V201505.ComposedLook
            

        
### Properties

#### Name

#### ColorFile

#### FontFile

#### BackgroundFile

#### MasterPage

#### SiteLogo

#### AlternateCSS

#### Version

#### VersionSpecified


## Core.Framework.Provisioning.Providers.Xml.V201505.Provider
            

        
### Properties

#### Configuration

#### Enabled

#### HandlerType


## Core.Framework.Provisioning.Providers.Xml.V201505.SiteCollection
            

        
### Properties

#### Templates

#### StorageMaximumLevel

#### StorageWarningLevel

#### UserCodeMaximumLevel

#### UserCodeWarningLevel

#### PrimarySiteCollectionAdmin

#### SecondarySiteCollectionAdmin

#### Title

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### Url


## Core.Framework.Provisioning.Providers.Xml.V201505.Sequence
            

        
### Fields

#### 

#### 

### Properties

#### SiteCollection

#### Site

#### TermStore

#### Extensions

#### SequenceType

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201505.SequenceSequenceType
            

        
### Fields

#### Synchronous

#### Asynchronous


## Core.Framework.Provisioning.Providers.Xml.V201508.Provisioning
            

        
### Properties

#### Preferences

#### Templates

#### Sequence

#### ImportSequence

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.Preferences
            

        
### Properties

#### Parameters

#### Version

#### Author

#### Generator

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.PreferencesParameter
            

        
### Properties

#### Key

#### Required

#### Text


## Core.Framework.Provisioning.Providers.Xml.V201508.ImportSequence
            

        
### Properties

#### File


## Core.Framework.Provisioning.Providers.Xml.V201508.TermStore
            

        
### Fields

#### 

#### 

### Properties

#### TermGroup

#### Scope


## Core.Framework.Provisioning.Providers.Xml.V201508.TermGroup
            

        
### Properties

#### TermSets

#### Description

#### SiteCollectionTermGroup

#### SiteCollectionTermGroupSpecified


## Core.Framework.Provisioning.Providers.Xml.V201508.TermSet
            

        
### Properties

#### CustomProperties

#### Terms

#### Language

#### LanguageSpecified

#### IsOpenForTermCreation

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.StringDictionaryItem
            

        
### Properties

#### Key

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201508.PropertyBagEntry
            

        
### Properties

#### Overwrite

#### OverwriteSpecified

#### Indexed


## Core.Framework.Provisioning.Providers.Xml.V201508.Term
            

        
### Fields

#### 

#### 

### Properties

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### Terms

#### Labels

#### CustomProperties

#### LocalCustomProperties

#### Language

#### LanguageSpecified

#### CustomSortOrder

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.TermTerms
            

        
### Properties

#### Items


## Core.Framework.Provisioning.Providers.Xml.V201508.TermLabelsLabel
            

        
### Properties

#### Language

#### Value

#### IsDefaultForLanguage


## Core.Framework.Provisioning.Providers.Xml.V201508.TermSetItem
            

        
### Properties

#### Owner

#### Description

#### IsAvailableForTagging


## Core.Framework.Provisioning.Providers.Xml.V201508.TaxonomyItem
            

        
### Properties

#### Name

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201508.TermStoreScope
            

        
### Fields

#### Default

#### Current


## Core.Framework.Provisioning.Providers.Xml.V201508.Site
            

        
### Properties

#### Templates

#### Title

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### UseSamePermissionsAsParentSite

#### UseSamePermissionsAsParentSiteSpecified

#### Url

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.Templates
            

        
### Properties

#### ProvisioningTemplateFile

#### ProvisioningTemplateReference

#### ProvisioningTemplate

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201508.ProvisioningTemplateFile
            

        
### Properties

#### File

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201508.ProvisioningTemplateReference
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201508.ProvisioningTemplate
            

        
### Properties

#### 

#### 

#### 

#### Properties

#### SitePolicy

#### RegionalSettings

#### SupportedUILanguages

#### AuditSettings

#### PropertyBagEntries

#### Security

#### SiteFields

#### ContentTypes

#### Lists

#### Features

#### CustomActions

#### Files

#### Pages

#### TermGroups

#### ComposedLook

#### Workflows

#### SearchSettings

#### Publishing

#### AddIns

#### Providers

#### ID

#### Version

#### VersionSpecified

#### ImagePreviewUrl

#### DisplayName

#### Description

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.RegionalSettings
            

        
### Properties

#### AdjustHijriDays

#### AdjustHijriDaysSpecified

#### AlternateCalendarType

#### AlternateCalendarTypeSpecified

#### CalendarType

#### CalendarTypeSpecified

#### Collation

#### CollationSpecified

#### FirstDayOfWeek

#### FirstDayOfWeekSpecified

#### FirstWeekOfYear

#### FirstWeekOfYearSpecified

#### LocaleId

#### LocaleIdSpecified

#### ShowWeeks

#### ShowWeeksSpecified

#### Time24

#### Time24Specified

#### TimeZone

#### WorkDayEndHour

#### WorkDayEndHourSpecified

#### WorkDays

#### WorkDaysSpecified

#### WorkDayStartHour

#### WorkDayStartHourSpecified


## Core.Framework.Provisioning.Providers.Xml.V201508.CalendarType
            

        
### Fields

#### None

#### Gregorian

#### Japan

#### Taiwan

#### Korea

#### Hijri

#### Thai

#### Hebrew

#### GregorianMiddleEastFrenchCalendar

#### GregorianArabicCalendar

#### GregorianTransliteratedEnglishCalendar

#### GregorianTransliteratedFrenchCalendar

#### KoreaandJapaneseLunar

#### ChineseLunar

#### SakaEra

#### UmmalQura


## Core.Framework.Provisioning.Providers.Xml.V201508.DayOfWeek
            

        
### Fields

#### Sunday

#### Monday

#### Tuesday

#### Wednesday

#### Thursday

#### Friday

#### Saturday


## Core.Framework.Provisioning.Providers.Xml.V201508.WorkHour
            

        
### Fields

#### Item1200AM

#### Item100AM

#### Item200AM

#### Item300AM

#### Item400AM

#### Item500AM

#### Item600AM

#### Item700AM

#### Item800AM

#### Item900AM

#### Item1000AM

#### Item1100AM

#### Item1200PM

#### Item100PM

#### Item200PM

#### Item300PM

#### Item400PM

#### Item500PM

#### Item600PM

#### Item700PM

#### Item800PM

#### Item900PM

#### Item1000PM

#### Item1100PM


## Core.Framework.Provisioning.Providers.Xml.V201508.SupportedUILanguagesSupportedUILanguage
            

        
### Properties

#### LCID


## Core.Framework.Provisioning.Providers.Xml.V201508.AuditSettings
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### Audit

#### AuditLogTrimmingRetention

#### AuditLogTrimmingRetentionSpecified

#### TrimAuditLog

#### TrimAuditLogSpecified

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.AuditSettingsAudit
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### AuditFlag


## Core.Framework.Provisioning.Providers.Xml.V201508.AuditSettingsAuditAuditFlag
            

        
### Fields

#### All

#### CheckIn

#### CheckOut

#### ChildDelete

#### Copy

#### Move

#### None

#### ObjectDelete

#### ProfileChange

#### SchemaChange

#### Search

#### SecurityChange

#### Undelete

#### Update

#### View

#### Workflow


## Core.Framework.Provisioning.Providers.Xml.V201508.Security
            

        
### Properties

#### AdditionalAdministrators

#### AdditionalOwners

#### AdditionalMembers

#### AdditionalVisitors

#### SiteGroups

#### Permissions

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.User
            

        
### Properties

#### Name


## Core.Framework.Provisioning.Providers.Xml.V201508.SiteGroup
            

        
### Properties

#### Members

#### Title

#### Description

#### Owner

#### AllowMembersEditMembership

#### AllowMembersEditMembershipSpecified

#### AllowRequestToJoinLeave

#### AllowRequestToJoinLeaveSpecified

#### AutoAcceptRequestToJoinLeave

#### AutoAcceptRequestToJoinLeaveSpecified

#### OnlyAllowMembersViewMembership

#### OnlyAllowMembersViewMembershipSpecified

#### RequestToJoinLeaveEmailSetting


## Core.Framework.Provisioning.Providers.Xml.V201508.SecurityPermissions
            

        
### Properties

#### RoleDefinitions

#### RoleAssignments


## Core.Framework.Provisioning.Providers.Xml.V201508.RoleDefinition
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### Permissions

#### Name

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201508.RoleDefinitionPermission
            

        
### Fields

#### EmptyMask

#### ViewListItems

#### AddListItems

#### EditListItems

#### DeleteListItems

#### ApproveItems

#### OpenItems

#### ViewVersions

#### DeleteVersions

#### CancelCheckout

#### ManagePersonalViews

#### ManageLists

#### ViewFormPages

#### AnonymousSearchAccessList

#### Open

#### ViewPages

#### AddAndCustomizePages

#### ApplyThemeAndBorder

#### ApplyStyleSheets

#### ViewUsageData

#### CreateSSCSite

#### ManageSubwebs

#### CreateGroups

#### ManagePermissions

#### BrowseDirectories

#### BrowseUserInfo

#### AddDelPrivateWebParts

#### UpdatePersonalWebParts

#### ManageWeb

#### AnonymousSearchAccessWebLists

#### UseClientIntegration

#### UseRemoteAPIs

#### ManageAlerts

#### CreateAlerts

#### EditMyUserInfo

#### EnumeratePermissions

#### FullMask


## Core.Framework.Provisioning.Providers.Xml.V201508.RoleAssignment
            

        
### Properties

#### Principal

#### RoleDefinition


## Core.Framework.Provisioning.Providers.Xml.V201508.ProvisioningTemplateSiteFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201508.ContentType
            

        
### Properties

#### FieldRefs

#### DocumentTemplate

#### DocumentSetTemplate

#### ID

#### Name

#### Description

#### Group

#### Hidden

#### Sealed

#### ReadOnly

#### Overwrite

#### NewFormUrl

#### EditFormUrl

#### DisplayFormUrl

#### AnyAttr

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.ContentTypeFieldRef
            

        

## Core.Framework.Provisioning.Providers.Xml.V201508.FieldRefFull
            

        
### Properties

#### Name

#### Required

#### Hidden


## Core.Framework.Provisioning.Providers.Xml.V201508.FieldRefBase
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201508.DocumentSetFieldRef
            

        

## Core.Framework.Provisioning.Providers.Xml.V201508.ListInstanceFieldRef
            

        
### Properties

#### DisplayName


## Core.Framework.Provisioning.Providers.Xml.V201508.ContentTypeDocumentTemplate
            

        
### Properties

#### TargetName


## Core.Framework.Provisioning.Providers.Xml.V201508.DocumentSetTemplate
            

        
### Properties

#### AllowedContentTypes

#### DefaultDocuments

#### SharedFields

#### WelcomePageFields

#### WelcomePage

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.DocumentSetTemplateAllowedContentType
            

        
### Properties

#### ContentTypeID


## Core.Framework.Provisioning.Providers.Xml.V201508.DocumentSetTemplateDefaultDocument
            

        
### Properties

#### Name

#### ContentTypeID

#### FileSourcePath


## Core.Framework.Provisioning.Providers.Xml.V201508.ListInstance
            

        
### Properties

#### 

#### ContentTypeBindings

#### Views

#### Fields

#### FieldRefs

#### DataRows

#### FieldDefaults

#### Security

#### Title

#### Description

#### DocumentTemplate

#### OnQuickLaunch

#### TemplateType

#### Url

#### EnableVersioning

#### EnableMinorVersions

#### EnableModeration

#### MinorVersionLimit

#### MinorVersionLimitSpecified

#### MaxVersionLimit

#### MaxVersionLimitSpecified

#### DraftVersionVisibility

#### DraftVersionVisibilitySpecified

#### RemoveExistingContentTypes

#### TemplateFeatureID

#### ContentTypesEnabled

#### Hidden

#### EnableAttachments

#### EnableFolderCreation

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.ContentTypeBinding
            

        
### Properties

#### ContentTypeID

#### Default


## Core.Framework.Provisioning.Providers.Xml.V201508.ListInstanceViews
            

        
### Properties

#### Any

#### RemoveExistingViews


## Core.Framework.Provisioning.Providers.Xml.V201508.ListInstanceFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201508.ListInstanceDataRow
            

        
### Properties

#### DataValue

#### Security


## Core.Framework.Provisioning.Providers.Xml.V201508.DataValue
            

        

## Core.Framework.Provisioning.Providers.Xml.V201508.BaseFieldValue
            

        
### Properties

#### FieldName

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201508.FieldDefault
            

        

## Core.Framework.Provisioning.Providers.Xml.V201508.ObjectSecurity
            

        
### Properties

#### BreakRoleInheritance

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.ObjectSecurityBreakRoleInheritance
            

        
### Properties

#### RoleAssignment

#### CopyRoleAssignments

#### ClearSubscopes


## Core.Framework.Provisioning.Providers.Xml.V201508.Features
            

        
### Properties

#### SiteFeatures

#### WebFeatures


## Core.Framework.Provisioning.Providers.Xml.V201508.Feature
            

        
### Properties

#### 

#### 

#### ID

#### Deactivate

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201508.CustomActions
            

        
### Properties

#### SiteCustomActions

#### WebCustomActions


## Core.Framework.Provisioning.Providers.Xml.V201508.CustomAction
            

        
### Properties

#### 

#### 

#### CommandUIExtension

#### Name

#### Description

#### Group

#### Location

#### Title

#### Sequence

#### SequenceSpecified

#### Rights

#### RightsSpecified

#### Url

#### Enabled

#### ScriptBlock

#### ImageUrl

#### ScriptSrc

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.CustomActionCommandUIExtension
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201508.File
            

        
### Properties

#### Properties

#### WebParts

#### Security

#### Src

#### Folder

#### Overwrite


## Core.Framework.Provisioning.Providers.Xml.V201508.WebPartPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Zone

#### Order


## Core.Framework.Provisioning.Providers.Xml.V201508.Pages
            

        
### Properties

#### Page

#### WelcomePage


## Core.Framework.Provisioning.Providers.Xml.V201508.Page
            

        
### Properties

#### 

#### 

#### WebParts

#### Security

#### Url

#### Overwrite

#### Layout


## Core.Framework.Provisioning.Providers.Xml.V201508.WikiPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Row

#### Column


## Core.Framework.Provisioning.Providers.Xml.V201508.WikiPageLayout
            

        
### Fields

#### OneColumn

#### OneColumnSidebar

#### TwoColumns

#### TwoColumnsHeader

#### TwoColumnsHeaderFooter

#### ThreeColumns

#### ThreeColumnsHeader

#### ThreeColumnsHeaderFooter


## Core.Framework.Provisioning.Providers.Xml.V201508.ComposedLook
            

        
### Properties

#### Name

#### ColorFile

#### FontFile

#### BackgroundFile

#### MasterPage

#### SiteLogo

#### AlternateCSS

#### Version

#### VersionSpecified


## Core.Framework.Provisioning.Providers.Xml.V201508.Workflows
            

        
### Fields

#### 

#### 

#### 

### Properties

#### WorkflowDefinitions

#### WorkflowSubscriptions

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.WorkflowsWorkflowDefinition
            

        
### Fields

#### 

#### 

#### 

### Properties

#### Properties

#### FormField

#### Id

#### AssociationUrl

#### Description

#### DisplayName

#### DraftVersion

#### InitiationUrl

#### Published

#### PublishedSpecified

#### RequiresAssociationForm

#### RequiresAssociationFormSpecified

#### RequiresInitiationForm

#### RequiresInitiationFormSpecified

#### RestrictToScope

#### RestrictToType

#### RestrictToTypeSpecified

#### XamlPath


## Core.Framework.Provisioning.Providers.Xml.V201508.WorkflowsWorkflowDefinitionRestrictToType
            

        
### Fields

#### Universal

#### List

#### Site


## Core.Framework.Provisioning.Providers.Xml.V201508.WorkflowsWorkflowSubscription
            

        
### Properties

#### PropertyDefinitions

#### DefinitionId

#### ListId

#### Enabled

#### EventSourceId

#### WorkflowStartEvent

#### ItemAddedEvent

#### ItemUpdatedEvent

#### ManualStartBypassesActivationLimit

#### ManualStartBypassesActivationLimitSpecified

#### Name

#### ParentContentTypeId

#### StatusFieldName


## Core.Framework.Provisioning.Providers.Xml.V201508.Publishing
            

        
### Fields

#### 

#### 

#### 

### Properties

#### DesignPackage

#### AvailableWebTemplates

#### PageLayouts

#### AutoCheckRequirements

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.PublishingDesignPackage
            

        
### Properties

#### DesignPackagePath

#### MajorVersion

#### MajorVersionSpecified

#### MinorVersion

#### MinorVersionSpecified

#### PackageGuid

#### PackageName


## Core.Framework.Provisioning.Providers.Xml.V201508.PublishingWebTemplate
            

        
### Properties

#### LanguageCode

#### LanguageCodeSpecified

#### TemplateName


## Core.Framework.Provisioning.Providers.Xml.V201508.PublishingPageLayouts
            

        
### Properties

#### PageLayout

#### Default

#### 


## Core.Framework.Provisioning.Providers.Xml.V201508.PublishingPageLayoutsPageLayout
            

        
### Properties

#### Path


## Core.Framework.Provisioning.Providers.Xml.V201508.PublishingAutoCheckRequirements
            

        
### Fields

#### MakeCompliant

#### SkipIfNotCompliant

#### FailIfNotCompliant


## Core.Framework.Provisioning.Providers.Xml.V201508.AddInsAddin
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### PackagePath

#### Source


## Core.Framework.Provisioning.Providers.Xml.V201508.AddInsAddinSource
            

        
### Fields

#### CorporateCatalog

#### DeveloperSite

#### InvalidSource

#### Marketplace

#### ObjectModel

#### RemoteObjectModel


## Core.Framework.Provisioning.Providers.Xml.V201508.Provider
            

        
### Properties

#### Configuration

#### Enabled

#### HandlerType


## Core.Framework.Provisioning.Providers.Xml.V201508.SiteCollection
            

        
### Properties

#### Templates

#### StorageMaximumLevel

#### StorageWarningLevel

#### UserCodeMaximumLevel

#### UserCodeWarningLevel

#### PrimarySiteCollectionAdmin

#### SecondarySiteCollectionAdmin

#### Title

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### Url


## Core.Framework.Provisioning.Providers.Xml.V201508.Sequence
            

        
### Fields

#### 

#### 

### Properties

#### SiteCollection

#### Site

#### TermStore

#### Extensions

#### SequenceType

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201508.SequenceSequenceType
            

        
### Fields

#### Synchronous

#### Asynchronous


## Core.Framework.Provisioning.Providers.Xml.V201512.Provisioning
            

        
### Properties

#### Preferences

#### Localizations

#### Templates

#### Sequence

#### ImportSequence

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.Preferences
            

        
### Properties

#### Parameters

#### Version

#### Author

#### Generator

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.PreferencesParameter
            

        
### Properties

#### Key

#### Required

#### Text


## Core.Framework.Provisioning.Providers.Xml.V201512.ImportSequence
            

        
### Properties

#### File


## Core.Framework.Provisioning.Providers.Xml.V201512.TermStore
            

        
### Fields

#### 

#### 

### Properties

#### TermGroup

#### Scope


## Core.Framework.Provisioning.Providers.Xml.V201512.TermGroup
            

        
### Properties

#### TermSets

#### Description

#### SiteCollectionTermGroup

#### SiteCollectionTermGroupSpecified


## Core.Framework.Provisioning.Providers.Xml.V201512.TermSet
            

        
### Properties

#### CustomProperties

#### Terms

#### Language

#### LanguageSpecified

#### IsOpenForTermCreation

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.StringDictionaryItem
            

        
### Properties

#### Key

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201512.PropertyBagEntry
            

        
### Properties

#### Overwrite

#### OverwriteSpecified

#### Indexed


## Core.Framework.Provisioning.Providers.Xml.V201512.Term
            

        
### Fields

#### 

#### 

### Properties

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### Terms

#### Labels

#### CustomProperties

#### LocalCustomProperties

#### Language

#### LanguageSpecified

#### CustomSortOrder

#### IsReused

#### IsSourceTerm

#### IsDeprecated

#### SourceTermId

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.TermTerms
            

        
### Properties

#### Items


## Core.Framework.Provisioning.Providers.Xml.V201512.TermLabelsLabel
            

        
### Properties

#### Language

#### Value

#### IsDefaultForLanguage


## Core.Framework.Provisioning.Providers.Xml.V201512.TermSetItem
            

        
### Properties

#### Owner

#### Description

#### IsAvailableForTagging


## Core.Framework.Provisioning.Providers.Xml.V201512.TaxonomyItem
            

        
### Properties

#### Name

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201512.TermStoreScope
            

        
### Fields

#### Default

#### Current


## Core.Framework.Provisioning.Providers.Xml.V201512.Site
            

        
### Properties

#### Templates

#### Title

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### UseSamePermissionsAsParentSite

#### UseSamePermissionsAsParentSiteSpecified

#### Url

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.Templates
            

        
### Properties

#### ProvisioningTemplateFile

#### ProvisioningTemplateReference

#### ProvisioningTemplate

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201512.ProvisioningTemplateFile
            

        
### Properties

#### File

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201512.ProvisioningTemplateReference
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201512.ProvisioningTemplate
            

        
### Properties

#### 

#### 

#### 

#### Properties

#### SitePolicy

#### WebSettings

#### RegionalSettings

#### SupportedUILanguages

#### AuditSettings

#### PropertyBagEntries

#### Security

#### SiteFields

#### ContentTypes

#### Lists

#### Features

#### CustomActions

#### Files

#### Pages

#### TermGroups

#### ComposedLook

#### Workflows

#### SearchSettings

#### Publishing

#### AddIns

#### Providers

#### ID

#### Version

#### VersionSpecified

#### ImagePreviewUrl

#### DisplayName

#### Description

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.WebSettings
            

        
### Properties

#### RequestAccessEmail

#### NoCrawl

#### NoCrawlSpecified

#### WelcomePage

#### Title

#### Description

#### SiteLogo

#### AlternateCSS

#### MasterPageUrl

#### CustomMasterPageUrl


## Core.Framework.Provisioning.Providers.Xml.V201512.RegionalSettings
            

        
### Properties

#### AdjustHijriDays

#### AdjustHijriDaysSpecified

#### AlternateCalendarType

#### AlternateCalendarTypeSpecified

#### CalendarType

#### CalendarTypeSpecified

#### Collation

#### CollationSpecified

#### FirstDayOfWeek

#### FirstDayOfWeekSpecified

#### FirstWeekOfYear

#### FirstWeekOfYearSpecified

#### LocaleId

#### LocaleIdSpecified

#### ShowWeeks

#### ShowWeeksSpecified

#### Time24

#### Time24Specified

#### TimeZone

#### WorkDayEndHour

#### WorkDayEndHourSpecified

#### WorkDays

#### WorkDaysSpecified

#### WorkDayStartHour

#### WorkDayStartHourSpecified


## Core.Framework.Provisioning.Providers.Xml.V201512.CalendarType
            

        
### Fields

#### None

#### Gregorian

#### Japan

#### Taiwan

#### Korea

#### Hijri

#### Thai

#### Hebrew

#### GregorianMiddleEastFrenchCalendar

#### GregorianArabicCalendar

#### GregorianTransliteratedEnglishCalendar

#### GregorianTransliteratedFrenchCalendar

#### KoreaandJapaneseLunar

#### ChineseLunar

#### SakaEra

#### UmmalQura


## Core.Framework.Provisioning.Providers.Xml.V201512.DayOfWeek
            

        
### Fields

#### Sunday

#### Monday

#### Tuesday

#### Wednesday

#### Thursday

#### Friday

#### Saturday


## Core.Framework.Provisioning.Providers.Xml.V201512.WorkHour
            

        
### Fields

#### Item1200AM

#### Item100AM

#### Item200AM

#### Item300AM

#### Item400AM

#### Item500AM

#### Item600AM

#### Item700AM

#### Item800AM

#### Item900AM

#### Item1000AM

#### Item1100AM

#### Item1200PM

#### Item100PM

#### Item200PM

#### Item300PM

#### Item400PM

#### Item500PM

#### Item600PM

#### Item700PM

#### Item800PM

#### Item900PM

#### Item1000PM

#### Item1100PM


## Core.Framework.Provisioning.Providers.Xml.V201512.SupportedUILanguagesSupportedUILanguage
            

        
### Properties

#### LCID


## Core.Framework.Provisioning.Providers.Xml.V201512.AuditSettings
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### Audit

#### AuditLogTrimmingRetention

#### AuditLogTrimmingRetentionSpecified

#### TrimAuditLog

#### TrimAuditLogSpecified

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.AuditSettingsAudit
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### AuditFlag


## Core.Framework.Provisioning.Providers.Xml.V201512.AuditSettingsAuditAuditFlag
            

        
### Fields

#### All

#### CheckIn

#### CheckOut

#### ChildDelete

#### Copy

#### Move

#### None

#### ObjectDelete

#### ProfileChange

#### SchemaChange

#### Search

#### SecurityChange

#### Undelete

#### Update

#### View

#### Workflow


## Core.Framework.Provisioning.Providers.Xml.V201512.Security
            

        
### Properties

#### AdditionalAdministrators

#### AdditionalOwners

#### AdditionalMembers

#### AdditionalVisitors

#### SiteGroups

#### Permissions

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.User
            

        
### Properties

#### Name


## Core.Framework.Provisioning.Providers.Xml.V201512.SiteGroup
            

        
### Properties

#### Members

#### Title

#### Description

#### Owner

#### AllowMembersEditMembership

#### AllowMembersEditMembershipSpecified

#### AllowRequestToJoinLeave

#### AllowRequestToJoinLeaveSpecified

#### AutoAcceptRequestToJoinLeave

#### AutoAcceptRequestToJoinLeaveSpecified

#### OnlyAllowMembersViewMembership

#### OnlyAllowMembersViewMembershipSpecified

#### RequestToJoinLeaveEmailSetting


## Core.Framework.Provisioning.Providers.Xml.V201512.SecurityPermissions
            

        
### Properties

#### RoleDefinitions

#### RoleAssignments


## Core.Framework.Provisioning.Providers.Xml.V201512.RoleDefinition
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### Permissions

#### Name

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201512.RoleDefinitionPermission
            

        
### Fields

#### EmptyMask

#### ViewListItems

#### AddListItems

#### EditListItems

#### DeleteListItems

#### ApproveItems

#### OpenItems

#### ViewVersions

#### DeleteVersions

#### CancelCheckout

#### ManagePersonalViews

#### ManageLists

#### ViewFormPages

#### AnonymousSearchAccessList

#### Open

#### ViewPages

#### AddAndCustomizePages

#### ApplyThemeAndBorder

#### ApplyStyleSheets

#### ViewUsageData

#### CreateSSCSite

#### ManageSubwebs

#### CreateGroups

#### ManagePermissions

#### BrowseDirectories

#### BrowseUserInfo

#### AddDelPrivateWebParts

#### UpdatePersonalWebParts

#### ManageWeb

#### AnonymousSearchAccessWebLists

#### UseClientIntegration

#### UseRemoteAPIs

#### ManageAlerts

#### CreateAlerts

#### EditMyUserInfo

#### EnumeratePermissions

#### FullMask


## Core.Framework.Provisioning.Providers.Xml.V201512.RoleAssignment
            

        
### Properties

#### Principal

#### RoleDefinition


## Core.Framework.Provisioning.Providers.Xml.V201512.ProvisioningTemplateSiteFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201512.ContentType
            

        
### Properties

#### FieldRefs

#### DocumentTemplate

#### DocumentSetTemplate

#### ID

#### Name

#### Description

#### Group

#### Hidden

#### Sealed

#### ReadOnly

#### Overwrite

#### NewFormUrl

#### EditFormUrl

#### DisplayFormUrl

#### AnyAttr

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.ContentTypeFieldRef
            

        

## Core.Framework.Provisioning.Providers.Xml.V201512.FieldRefFull
            

        
### Properties

#### Name

#### Required

#### Hidden


## Core.Framework.Provisioning.Providers.Xml.V201512.FieldRefBase
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201512.DocumentSetFieldRef
            

        

## Core.Framework.Provisioning.Providers.Xml.V201512.ListInstanceFieldRef
            

        
### Properties

#### DisplayName


## Core.Framework.Provisioning.Providers.Xml.V201512.ContentTypeDocumentTemplate
            

        
### Properties

#### TargetName


## Core.Framework.Provisioning.Providers.Xml.V201512.DocumentSetTemplate
            

        
### Properties

#### AllowedContentTypes

#### DefaultDocuments

#### SharedFields

#### WelcomePageFields

#### WelcomePage

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.DocumentSetTemplateAllowedContentType
            

        
### Properties

#### ContentTypeID


## Core.Framework.Provisioning.Providers.Xml.V201512.DocumentSetTemplateDefaultDocument
            

        
### Properties

#### Name

#### ContentTypeID

#### FileSourcePath


## Core.Framework.Provisioning.Providers.Xml.V201512.ListInstance
            

        
### Properties

#### 

#### ContentTypeBindings

#### Views

#### Fields

#### FieldRefs

#### DataRows

#### Folders

#### FieldDefaults

#### Security

#### Title

#### Description

#### DocumentTemplate

#### OnQuickLaunch

#### TemplateType

#### Url

#### EnableVersioning

#### EnableMinorVersions

#### EnableModeration

#### MinorVersionLimit

#### MinorVersionLimitSpecified

#### MaxVersionLimit

#### MaxVersionLimitSpecified

#### DraftVersionVisibility

#### DraftVersionVisibilitySpecified

#### RemoveExistingContentTypes

#### TemplateFeatureID

#### ContentTypesEnabled

#### Hidden

#### EnableAttachments

#### EnableFolderCreation

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.ContentTypeBinding
            

        
### Properties

#### ContentTypeID

#### Default


## Core.Framework.Provisioning.Providers.Xml.V201512.ListInstanceViews
            

        
### Properties

#### Any

#### RemoveExistingViews


## Core.Framework.Provisioning.Providers.Xml.V201512.ListInstanceFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201512.ListInstanceDataRow
            

        
### Properties

#### DataValue

#### Security


## Core.Framework.Provisioning.Providers.Xml.V201512.DataValue
            

        

## Core.Framework.Provisioning.Providers.Xml.V201512.BaseFieldValue
            

        
### Properties

#### FieldName

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201512.FieldDefault
            

        

## Core.Framework.Provisioning.Providers.Xml.V201512.ObjectSecurity
            

        
### Properties

#### BreakRoleInheritance

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.ObjectSecurityBreakRoleInheritance
            

        
### Properties

#### RoleAssignment

#### CopyRoleAssignments

#### ClearSubscopes


## Core.Framework.Provisioning.Providers.Xml.V201512.Folder
            

        
### Properties

#### Folder1

#### Security

#### Name


## Core.Framework.Provisioning.Providers.Xml.V201512.Features
            

        
### Properties

#### SiteFeatures

#### WebFeatures


## Core.Framework.Provisioning.Providers.Xml.V201512.Feature
            

        
### Properties

#### 

#### 

#### ID

#### Deactivate

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201512.CustomActions
            

        
### Properties

#### SiteCustomActions

#### WebCustomActions


## Core.Framework.Provisioning.Providers.Xml.V201512.CustomAction
            

        
### Properties

#### 

#### 

#### CommandUIExtension

#### Name

#### Description

#### Group

#### Location

#### Title

#### Sequence

#### SequenceSpecified

#### Rights

#### Url

#### Enabled

#### ScriptBlock

#### ImageUrl

#### ScriptSrc

#### RegistrationId

#### RegistrationType

#### RegistrationTypeSpecified

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.CustomActionCommandUIExtension
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201512.RegistrationType
            

        
### Fields

#### None

#### List

#### ContentType

#### ProgId

#### FileType


## Core.Framework.Provisioning.Providers.Xml.V201512.File
            

        
### Properties

#### Properties

#### WebParts

#### Security

#### Src

#### Folder

#### Overwrite


## Core.Framework.Provisioning.Providers.Xml.V201512.WebPartPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Zone

#### Order


## Core.Framework.Provisioning.Providers.Xml.V201512.Page
            

        
### Properties

#### WebParts

#### Fields

#### Security

#### Url

#### Overwrite

#### Layout


## Core.Framework.Provisioning.Providers.Xml.V201512.WikiPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Row

#### Column


## Core.Framework.Provisioning.Providers.Xml.V201512.WikiPageLayout
            

        
### Fields

#### OneColumn

#### OneColumnSidebar

#### TwoColumns

#### TwoColumnsHeader

#### TwoColumnsHeaderFooter

#### ThreeColumns

#### ThreeColumnsHeader

#### ThreeColumnsHeaderFooter

#### Custom


## Core.Framework.Provisioning.Providers.Xml.V201512.ComposedLook
            

        
### Properties

#### Name

#### ColorFile

#### FontFile

#### BackgroundFile

#### Version

#### VersionSpecified


## Core.Framework.Provisioning.Providers.Xml.V201512.Workflows
            

        
### Fields

#### 

#### 

#### 

### Properties

#### WorkflowDefinitions

#### WorkflowSubscriptions

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.WorkflowsWorkflowDefinition
            

        
### Fields

#### 

#### 

#### 

### Properties

#### Properties

#### FormField

#### Id

#### AssociationUrl

#### Description

#### DisplayName

#### DraftVersion

#### InitiationUrl

#### Published

#### PublishedSpecified

#### RequiresAssociationForm

#### RequiresAssociationFormSpecified

#### RequiresInitiationForm

#### RequiresInitiationFormSpecified

#### RestrictToScope

#### RestrictToType

#### RestrictToTypeSpecified

#### XamlPath


## Core.Framework.Provisioning.Providers.Xml.V201512.WorkflowsWorkflowDefinitionRestrictToType
            

        
### Fields

#### Universal

#### List

#### Site


## Core.Framework.Provisioning.Providers.Xml.V201512.WorkflowsWorkflowSubscription
            

        
### Properties

#### PropertyDefinitions

#### DefinitionId

#### ListId

#### Enabled

#### EventSourceId

#### WorkflowStartEvent

#### ItemAddedEvent

#### ItemUpdatedEvent

#### ManualStartBypassesActivationLimit

#### ManualStartBypassesActivationLimitSpecified

#### Name

#### ParentContentTypeId

#### StatusFieldName


## Core.Framework.Provisioning.Providers.Xml.V201512.Publishing
            

        
### Fields

#### 

#### 

#### 

### Properties

#### DesignPackage

#### AvailableWebTemplates

#### PageLayouts

#### AutoCheckRequirements

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.PublishingDesignPackage
            

        
### Properties

#### DesignPackagePath

#### MajorVersion

#### MajorVersionSpecified

#### MinorVersion

#### MinorVersionSpecified

#### PackageGuid

#### PackageName


## Core.Framework.Provisioning.Providers.Xml.V201512.PublishingWebTemplate
            

        
### Properties

#### LanguageCode

#### LanguageCodeSpecified

#### TemplateName


## Core.Framework.Provisioning.Providers.Xml.V201512.PublishingPageLayouts
            

        
### Properties

#### PageLayout

#### Default

#### 


## Core.Framework.Provisioning.Providers.Xml.V201512.PublishingPageLayoutsPageLayout
            

        
### Properties

#### Path


## Core.Framework.Provisioning.Providers.Xml.V201512.PublishingAutoCheckRequirements
            

        
### Fields

#### MakeCompliant

#### SkipIfNotCompliant

#### FailIfNotCompliant


## Core.Framework.Provisioning.Providers.Xml.V201512.AddInsAddin
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### PackagePath

#### Source


## Core.Framework.Provisioning.Providers.Xml.V201512.AddInsAddinSource
            

        
### Fields

#### CorporateCatalog

#### DeveloperSite

#### InvalidSource

#### Marketplace

#### ObjectModel

#### RemoteObjectModel


## Core.Framework.Provisioning.Providers.Xml.V201512.Provider
            

        
### Properties

#### Configuration

#### Enabled

#### HandlerType


## Core.Framework.Provisioning.Providers.Xml.V201512.SiteCollection
            

        
### Properties

#### Templates

#### StorageMaximumLevel

#### StorageWarningLevel

#### UserCodeMaximumLevel

#### UserCodeWarningLevel

#### PrimarySiteCollectionAdmin

#### SecondarySiteCollectionAdmin

#### Title

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### Url


## Core.Framework.Provisioning.Providers.Xml.V201512.Sequence
            

        
### Fields

#### 

#### 

### Properties

#### SiteCollection

#### Site

#### TermStore

#### Extensions

#### SequenceType

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201512.SequenceSequenceType
            

        
### Fields

#### Synchronous

#### Asynchronous


## Core.Framework.Provisioning.Providers.Xml.V201512.LocalizationsLocalization
            

        
### Properties

#### LCID

#### Name

#### ResourceFile


## Core.Framework.Provisioning.Providers.Xml.V201605.WikiPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Row

#### Column


## Core.Framework.Provisioning.Providers.Xml.V201605.BaseFieldValue
            

        
### Properties

#### FieldName

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201605.WebPartPageWebPart
            

        
### Properties

#### Contents

#### Title

#### Zone

#### Order


## Core.Framework.Provisioning.Providers.Xml.V201605.NavigationNode
            

        
### Properties

#### NavigationNode1

#### Title

#### Url

#### IsExternal

#### IsVisible


## Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning
            

        
### Properties

#### Preferences

#### Localizations

#### Templates

#### Sequence

#### ImportSequence

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.Preferences
            

        
### Properties

#### Parameters

#### Version

#### Author

#### Generator

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.PreferencesParameter
            

        
### Properties

#### Key

#### Required

#### Text


## Core.Framework.Provisioning.Providers.Xml.V201605.ImportSequence
            

        
### Properties

#### File


## Core.Framework.Provisioning.Providers.Xml.V201605.TermStore
            

        
### Fields

#### 

#### 

### Properties

#### TermGroup

#### Scope


## Core.Framework.Provisioning.Providers.Xml.V201605.TermGroup
            

        
### Properties

#### TermSets

#### Contributors

#### Managers

#### Description

#### SiteCollectionTermGroup

#### SiteCollectionTermGroupSpecified


## Core.Framework.Provisioning.Providers.Xml.V201605.TermSet
            

        
### Properties

#### CustomProperties

#### Terms

#### Language

#### LanguageSpecified

#### IsOpenForTermCreation

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.StringDictionaryItem
            

        
### Properties

#### Key

#### Value


## Core.Framework.Provisioning.Providers.Xml.V201605.PropertyBagEntry
            

        
### Properties

#### Overwrite

#### OverwriteSpecified

#### Indexed


## Core.Framework.Provisioning.Providers.Xml.V201605.Term
            

        
### Fields

#### 

#### 

### Properties

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### Terms

#### Labels

#### CustomProperties

#### LocalCustomProperties

#### Language

#### LanguageSpecified

#### CustomSortOrder

#### IsReused

#### IsSourceTerm

#### IsDeprecated

#### SourceTermId

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.TermTerms
            

        
### Properties

#### Items


## Core.Framework.Provisioning.Providers.Xml.V201605.TermLabelsLabel
            

        
### Properties

#### Language

#### Value

#### IsDefaultForLanguage


## Core.Framework.Provisioning.Providers.Xml.V201605.TermSetItem
            

        
### Properties

#### Owner

#### Description

#### IsAvailableForTagging


## Core.Framework.Provisioning.Providers.Xml.V201605.TaxonomyItem
            

        
### Properties

#### Name

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201605.User
            

        
### Properties

#### Name


## Core.Framework.Provisioning.Providers.Xml.V201605.TermStoreScope
            

        
### Fields

#### Default

#### Current


## Core.Framework.Provisioning.Providers.Xml.V201605.Site
            

        
### Properties

#### Templates

#### Title

#### SiteTemplate

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### UseSamePermissionsAsParentSite

#### UseSamePermissionsAsParentSiteSpecified

#### Url

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.Templates
            

        
### Properties

#### ProvisioningTemplateFile

#### ProvisioningTemplateReference

#### ProvisioningTemplate

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201605.ProvisioningTemplateFile
            

        
### Properties

#### File

#### ID

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.ProvisioningTemplateReference
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201605.ProvisioningTemplate
            

        
### Properties

#### 

#### 

#### 

#### Properties

#### SitePolicy

#### WebSettings

#### RegionalSettings

#### SupportedUILanguages

#### AuditSettings

#### PropertyBagEntries

#### Security

#### Navigation

#### SiteFields

#### ContentTypes

#### Lists

#### Features

#### CustomActions

#### Files

#### Pages

#### TermGroups

#### ComposedLook

#### Workflows

#### SearchSettings

#### Publishing

#### AddIns

#### Providers

#### ID

#### Version

#### VersionSpecified

#### BaseSiteTemplate

#### ImagePreviewUrl

#### DisplayName

#### Description

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.WebSettings
            

        
### Properties

#### RequestAccessEmail

#### NoCrawl

#### NoCrawlSpecified

#### WelcomePage

#### Title

#### Description

#### SiteLogo

#### AlternateCSS

#### MasterPageUrl

#### CustomMasterPageUrl


## Core.Framework.Provisioning.Providers.Xml.V201605.RegionalSettings
            

        
### Properties

#### AdjustHijriDays

#### AdjustHijriDaysSpecified

#### AlternateCalendarType

#### AlternateCalendarTypeSpecified

#### CalendarType

#### CalendarTypeSpecified

#### Collation

#### CollationSpecified

#### FirstDayOfWeek

#### FirstDayOfWeekSpecified

#### FirstWeekOfYear

#### FirstWeekOfYearSpecified

#### LocaleId

#### LocaleIdSpecified

#### ShowWeeks

#### ShowWeeksSpecified

#### Time24

#### Time24Specified

#### TimeZone

#### WorkDayEndHour

#### WorkDayEndHourSpecified

#### WorkDays

#### WorkDaysSpecified

#### WorkDayStartHour

#### WorkDayStartHourSpecified


## Core.Framework.Provisioning.Providers.Xml.V201605.CalendarType
            

        
### Fields

#### None

#### Gregorian

#### Japan

#### Taiwan

#### Korea

#### Hijri

#### Thai

#### Hebrew

#### GregorianMiddleEastFrenchCalendar

#### GregorianArabicCalendar

#### GregorianTransliteratedEnglishCalendar

#### GregorianTransliteratedFrenchCalendar

#### KoreaandJapaneseLunar

#### ChineseLunar

#### SakaEra

#### UmmalQura


## Core.Framework.Provisioning.Providers.Xml.V201605.DayOfWeek
            

        
### Fields

#### Sunday

#### Monday

#### Tuesday

#### Wednesday

#### Thursday

#### Friday

#### Saturday


## Core.Framework.Provisioning.Providers.Xml.V201605.WorkHour
            

        
### Fields

#### Item1200AM

#### Item100AM

#### Item200AM

#### Item300AM

#### Item400AM

#### Item500AM

#### Item600AM

#### Item700AM

#### Item800AM

#### Item900AM

#### Item1000AM

#### Item1100AM

#### Item1200PM

#### Item100PM

#### Item200PM

#### Item300PM

#### Item400PM

#### Item500PM

#### Item600PM

#### Item700PM

#### Item800PM

#### Item900PM

#### Item1000PM

#### Item1100PM


## Core.Framework.Provisioning.Providers.Xml.V201605.SupportedUILanguagesSupportedUILanguage
            

        
### Properties

#### LCID


## Core.Framework.Provisioning.Providers.Xml.V201605.AuditSettings
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### Audit

#### AuditLogTrimmingRetention

#### AuditLogTrimmingRetentionSpecified

#### TrimAuditLog

#### TrimAuditLogSpecified

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.AuditSettingsAudit
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### AuditFlag


## Core.Framework.Provisioning.Providers.Xml.V201605.AuditSettingsAuditAuditFlag
            

        
### Fields

#### All

#### CheckIn

#### CheckOut

#### ChildDelete

#### Copy

#### Move

#### None

#### ObjectDelete

#### ProfileChange

#### SchemaChange

#### Search

#### SecurityChange

#### Undelete

#### Update

#### View

#### Workflow


## Core.Framework.Provisioning.Providers.Xml.V201605.Security
            

        
### Properties

#### AdditionalAdministrators

#### AdditionalOwners

#### AdditionalMembers

#### AdditionalVisitors

#### SiteGroups

#### Permissions

#### BreakRoleInheritance

#### CopyRoleAssignments

#### ClearSubscopes

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.SiteGroup
            

        
### Properties

#### Members

#### Title

#### Description

#### Owner

#### AllowMembersEditMembership

#### AllowMembersEditMembershipSpecified

#### AllowRequestToJoinLeave

#### AllowRequestToJoinLeaveSpecified

#### AutoAcceptRequestToJoinLeave

#### AutoAcceptRequestToJoinLeaveSpecified

#### OnlyAllowMembersViewMembership

#### OnlyAllowMembersViewMembershipSpecified

#### RequestToJoinLeaveEmailSetting


## Core.Framework.Provisioning.Providers.Xml.V201605.SecurityPermissions
            

        
### Properties

#### RoleDefinitions

#### RoleAssignments


## Core.Framework.Provisioning.Providers.Xml.V201605.RoleDefinition
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### Permissions

#### Name

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201605.RoleDefinitionPermission
            

        
### Fields

#### EmptyMask

#### ViewListItems

#### AddListItems

#### EditListItems

#### DeleteListItems

#### ApproveItems

#### OpenItems

#### ViewVersions

#### DeleteVersions

#### CancelCheckout

#### ManagePersonalViews

#### ManageLists

#### ViewFormPages

#### AnonymousSearchAccessList

#### Open

#### ViewPages

#### AddAndCustomizePages

#### ApplyThemeAndBorder

#### ApplyStyleSheets

#### ViewUsageData

#### CreateSSCSite

#### ManageSubwebs

#### CreateGroups

#### ManagePermissions

#### BrowseDirectories

#### BrowseUserInfo

#### AddDelPrivateWebParts

#### UpdatePersonalWebParts

#### ManageWeb

#### AnonymousSearchAccessWebLists

#### UseClientIntegration

#### UseRemoteAPIs

#### ManageAlerts

#### CreateAlerts

#### EditMyUserInfo

#### EnumeratePermissions

#### FullMask


## Core.Framework.Provisioning.Providers.Xml.V201605.RoleAssignment
            

        
### Properties

#### Principal

#### RoleDefinition


## Core.Framework.Provisioning.Providers.Xml.V201605.Navigation
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### 

#### 

#### 

#### 

#### 

#### GlobalNavigation

#### CurrentNavigation

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.NavigationGlobalNavigation
            

        
### Fields

#### 

#### 

#### 

### Properties

#### StructuralNavigation

#### ManagedNavigation

#### NavigationType


## Core.Framework.Provisioning.Providers.Xml.V201605.StructuralNavigation
            

        
### Properties

#### NavigationNode

#### RemoveExistingNodes


## Core.Framework.Provisioning.Providers.Xml.V201605.ManagedNavigation
            

        
### Properties

#### TermStoreId

#### TermSetId


## Core.Framework.Provisioning.Providers.Xml.V201605.NavigationGlobalNavigationNavigationType
            

        
### Fields

#### Inherit

#### Structural

#### Managed


## Core.Framework.Provisioning.Providers.Xml.V201605.NavigationCurrentNavigation
            

        
### Fields

#### 

#### 

#### 

#### 

### Properties

#### StructuralNavigation

#### ManagedNavigation

#### NavigationType


## Core.Framework.Provisioning.Providers.Xml.V201605.NavigationCurrentNavigationNavigationType
            

        
### Fields

#### Inherit

#### Structural

#### StructuralLocal

#### Managed


## Core.Framework.Provisioning.Providers.Xml.V201605.ProvisioningTemplateSiteFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201605.ContentType
            

        
### Properties

#### FieldRefs

#### DocumentTemplate

#### DocumentSetTemplate

#### ID

#### Name

#### Description

#### Group

#### Hidden

#### Sealed

#### ReadOnly

#### Overwrite

#### NewFormUrl

#### EditFormUrl

#### DisplayFormUrl

#### AnyAttr

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.ContentTypeFieldRef
            

        

## Core.Framework.Provisioning.Providers.Xml.V201605.FieldRefFull
            

        
### Properties

#### Name

#### Required

#### Hidden


## Core.Framework.Provisioning.Providers.Xml.V201605.FieldRefBase
            

        
### Properties

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201605.DocumentSetFieldRef
            

        

## Core.Framework.Provisioning.Providers.Xml.V201605.ListInstanceFieldRef
            

        
### Properties

#### DisplayName


## Core.Framework.Provisioning.Providers.Xml.V201605.ContentTypeDocumentTemplate
            

        
### Properties

#### TargetName


## Core.Framework.Provisioning.Providers.Xml.V201605.DocumentSetTemplate
            

        
### Properties

#### AllowedContentTypes

#### DefaultDocuments

#### SharedFields

#### WelcomePageFields

#### WelcomePage

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.DocumentSetTemplateAllowedContentType
            

        
### Properties

#### ContentTypeID


## Core.Framework.Provisioning.Providers.Xml.V201605.DocumentSetTemplateDefaultDocument
            

        
### Properties

#### Name

#### ContentTypeID

#### FileSourcePath


## Core.Framework.Provisioning.Providers.Xml.V201605.ListInstance
            

        
### Properties

#### 

#### ContentTypeBindings

#### Views

#### Fields

#### FieldRefs

#### DataRows

#### Folders

#### FieldDefaults

#### Security

#### UserCustomActions

#### Title

#### Description

#### DocumentTemplate

#### OnQuickLaunch

#### TemplateType

#### Url

#### ForceCheckout

#### EnableVersioning

#### EnableMinorVersions

#### EnableModeration

#### MinorVersionLimit

#### MinorVersionLimitSpecified

#### MaxVersionLimit

#### MaxVersionLimitSpecified

#### DraftVersionVisibility

#### DraftVersionVisibilitySpecified

#### RemoveExistingContentTypes

#### TemplateFeatureID

#### ContentTypesEnabled

#### Hidden

#### EnableAttachments

#### EnableFolderCreation

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.ContentTypeBinding
            

        
### Properties

#### ContentTypeID

#### Default

#### Remove


## Core.Framework.Provisioning.Providers.Xml.V201605.ListInstanceViews
            

        
### Properties

#### Any

#### RemoveExistingViews


## Core.Framework.Provisioning.Providers.Xml.V201605.ListInstanceFields
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201605.ListInstanceDataRow
            

        
### Properties

#### DataValue

#### Security


## Core.Framework.Provisioning.Providers.Xml.V201605.DataValue
            

        

## Core.Framework.Provisioning.Providers.Xml.V201605.FieldDefault
            

        

## Core.Framework.Provisioning.Providers.Xml.V201605.ObjectSecurity
            

        
### Properties

#### BreakRoleInheritance

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.ObjectSecurityBreakRoleInheritance
            

        
### Properties

#### RoleAssignment

#### CopyRoleAssignments

#### ClearSubscopes


## Core.Framework.Provisioning.Providers.Xml.V201605.Folder
            

        
### Properties

#### Folder1

#### Security

#### Name


## Core.Framework.Provisioning.Providers.Xml.V201605.CustomAction
            

        
### Properties

#### CommandUIExtension

#### Name

#### Description

#### Group

#### Location

#### Title

#### Sequence

#### SequenceSpecified

#### Rights

#### Url

#### Enabled

#### Remove

#### ScriptBlock

#### ImageUrl

#### ScriptSrc

#### RegistrationId

#### RegistrationType

#### RegistrationTypeSpecified

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.CustomActionCommandUIExtension
            

        
### Properties

#### Any


## Core.Framework.Provisioning.Providers.Xml.V201605.RegistrationType
            

        
### Fields

#### None

#### List

#### ContentType

#### ProgId

#### FileType


## Core.Framework.Provisioning.Providers.Xml.V201605.Features
            

        
### Properties

#### SiteFeatures

#### WebFeatures


## Core.Framework.Provisioning.Providers.Xml.V201605.Feature
            

        
### Properties

#### 

#### 

#### ID

#### Deactivate

#### Description


## Core.Framework.Provisioning.Providers.Xml.V201605.CustomActions
            

        
### Properties

#### SiteCustomActions

#### WebCustomActions


## Core.Framework.Provisioning.Providers.Xml.V201605.ProvisioningTemplateFiles
            

        
### Properties

#### File

#### Directory


## Core.Framework.Provisioning.Providers.Xml.V201605.File
            

        
### Fields

#### 

#### 

#### 

### Properties

#### Properties

#### WebParts

#### Security

#### Src

#### Folder

#### Overwrite

#### Level

#### LevelSpecified


## Core.Framework.Provisioning.Providers.Xml.V201605.FileLevel
            

        
### Fields

#### Published

#### Draft

#### Checkout


## Core.Framework.Provisioning.Providers.Xml.V201605.Directory
            

        
### Properties

#### Security

#### Src

#### Folder

#### Overwrite

#### Level

#### LevelSpecified

#### Recursive

#### IncludedExtensions

#### ExcludedExtensions

#### MetadataMappingFile


## Core.Framework.Provisioning.Providers.Xml.V201605.Page
            

        
### Properties

#### WebParts

#### Fields

#### Security

#### Url

#### Overwrite

#### Layout


## Core.Framework.Provisioning.Providers.Xml.V201605.WikiPageLayout
            

        
### Fields

#### OneColumn

#### OneColumnSidebar

#### TwoColumns

#### TwoColumnsHeader

#### TwoColumnsHeaderFooter

#### ThreeColumns

#### ThreeColumnsHeader

#### ThreeColumnsHeaderFooter

#### Custom


## Core.Framework.Provisioning.Providers.Xml.V201605.ComposedLook
            

        
### Properties

#### Name

#### ColorFile

#### FontFile

#### BackgroundFile

#### Version

#### VersionSpecified


## Core.Framework.Provisioning.Providers.Xml.V201605.Workflows
            

        
### Fields

#### 

#### 

#### 

### Properties

#### WorkflowDefinitions

#### WorkflowSubscriptions

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.WorkflowsWorkflowDefinition
            

        
### Fields

#### 

#### 

#### 

### Properties

#### Properties

#### FormField

#### Id

#### AssociationUrl

#### Description

#### DisplayName

#### DraftVersion

#### InitiationUrl

#### Published

#### PublishedSpecified

#### RequiresAssociationForm

#### RequiresAssociationFormSpecified

#### RequiresInitiationForm

#### RequiresInitiationFormSpecified

#### RestrictToScope

#### RestrictToType

#### RestrictToTypeSpecified

#### XamlPath


## Core.Framework.Provisioning.Providers.Xml.V201605.WorkflowsWorkflowDefinitionRestrictToType
            

        
### Fields

#### Universal

#### List

#### Site


## Core.Framework.Provisioning.Providers.Xml.V201605.WorkflowsWorkflowSubscription
            

        
### Properties

#### PropertyDefinitions

#### DefinitionId

#### ListId

#### Enabled

#### EventSourceId

#### WorkflowStartEvent

#### ItemAddedEvent

#### ItemUpdatedEvent

#### ManualStartBypassesActivationLimit

#### ManualStartBypassesActivationLimitSpecified

#### Name

#### ParentContentTypeId

#### StatusFieldName


## Core.Framework.Provisioning.Providers.Xml.V201605.ProvisioningTemplateSearchSettings
            

        
### Properties

#### SiteSearchSettings

#### WebSearchSettings


## Core.Framework.Provisioning.Providers.Xml.V201605.Publishing
            

        
### Fields

#### 

#### 

#### 

### Properties

#### DesignPackage

#### AvailableWebTemplates

#### PageLayouts

#### AutoCheckRequirements

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.PublishingDesignPackage
            

        
### Properties

#### DesignPackagePath

#### MajorVersion

#### MajorVersionSpecified

#### MinorVersion

#### MinorVersionSpecified

#### PackageGuid

#### PackageName


## Core.Framework.Provisioning.Providers.Xml.V201605.PublishingWebTemplate
            

        
### Properties

#### LanguageCode

#### LanguageCodeSpecified

#### TemplateName


## Core.Framework.Provisioning.Providers.Xml.V201605.PublishingPageLayouts
            

        
### Properties

#### PageLayout

#### Default

#### 


## Core.Framework.Provisioning.Providers.Xml.V201605.PublishingPageLayoutsPageLayout
            

        
### Properties

#### Path


## Core.Framework.Provisioning.Providers.Xml.V201605.PublishingAutoCheckRequirements
            

        
### Fields

#### MakeCompliant

#### SkipIfNotCompliant

#### FailIfNotCompliant


## Core.Framework.Provisioning.Providers.Xml.V201605.AddInsAddin
            

        
### Fields

#### 

#### 

#### 

#### 

#### 

#### 

### Properties

#### PackagePath

#### Source


## Core.Framework.Provisioning.Providers.Xml.V201605.AddInsAddinSource
            

        
### Fields

#### CorporateCatalog

#### DeveloperSite

#### InvalidSource

#### Marketplace

#### ObjectModel

#### RemoteObjectModel


## Core.Framework.Provisioning.Providers.Xml.V201605.Provider
            

        
### Properties

#### Configuration

#### Enabled

#### HandlerType


## Core.Framework.Provisioning.Providers.Xml.V201605.SiteCollection
            

        
### Properties

#### Templates

#### StorageMaximumLevel

#### StorageWarningLevel

#### UserCodeMaximumLevel

#### UserCodeWarningLevel

#### PrimarySiteCollectionAdmin

#### SecondarySiteCollectionAdmin

#### Title

#### SiteTemplate

#### CustomJSUrl

#### QuickLaunchEnabled

#### QuickLaunchEnabledSpecified

#### AlternateCssUrl

#### Language

#### AllowDesigner

#### AllowDesignerSpecified

#### MembersCanShare

#### MembersCanShareSpecified

#### TimeZone

#### Url


## Core.Framework.Provisioning.Providers.Xml.V201605.Sequence
            

        
### Fields

#### 

#### 

### Properties

#### SiteCollection

#### Site

#### TermStore

#### Extensions

#### SequenceType

#### ID


## Core.Framework.Provisioning.Providers.Xml.V201605.SequenceSequenceType
            

        
### Fields

#### Synchronous

#### Asynchronous


## Core.Framework.Provisioning.Providers.Xml.V201605.LocalizationsLocalization
            

        
### Properties

#### LCID

#### Name

#### ResourceFile


## Core.Framework.Provisioning.Providers.Xml.TemplateSchemaSerializerAttribute
            
Attribute for Template Schema Serializers
        
### Properties

#### MinimalSupportedSchemaVersion
The schemas supported by the serializer
#### SerializationSequence
The sequence number for applying the serializer during serialization Should be a multiple of 100, to make room for future new insertions
#### DeserializationSequence
The sequence number for applying the serializer during deserialization Should be a multiple of 100, to make room for future new insertions
#### Default
Defines whether to automatically include the serializer in the serialization process or not

## Core.Framework.Provisioning.Providers.Xml.XMLPnPSchemaFormatter
            
Helper class that abstracts from any specific version of XMLPnPSchemaFormatter
        
### Properties

#### LatestFormatter
Static property to retrieve an instance of the latest XMLPnPSchemaFormatter
### Methods


#### GetSpecificFormatter(OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.XMLPnPSchemaVersion)
Static method to retrieve a specific XMLPnPSchemaFormatter instance
> ##### Parameters
> **version:** 

> ##### Return value
> 

#### GetSpecificFormatter(System.String)
Static method to retrieve a specific XMLPnPSchemaFormatter instance
> ##### Parameters
> **namespaceUri:** 

> ##### Return value
> 

## Core.Framework.Provisioning.Providers.Xml.XMLPnPSchemaV201605Serializer
            
Implements the logic to serialize a schema of version 201605
        

## Core.Framework.Provisioning.Providers.Xml.XMLTemplateProvider
            
Provider for xml based configurations
        

## Core.Framework.Provisioning.Providers.ITemplateFormatter
            
Interface for basic capabilites that any Template Formatter should provide/support
        
### Methods


#### Initialize(OfficeDevPnP.Core.Framework.Provisioning.Providers.TemplateProviderBase)
Method to initialize the formatter with the proper TemplateProvider instance
> ##### Parameters
> **provider:** The provider that is calling the current template formatter


#### IsValid(System.IO.Stream)
Method to validate the content of a formatted template instace
> ##### Parameters
> **template:** The formatted template instance as a Stream

> ##### Return value
> Boolean result of the validation

#### ToFormattedTemplate(OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate)
Method to format a ProvisioningTemplate into a formatted template
> ##### Parameters
> **template:** The input ProvisioningTemplate

> ##### Return value
> The output formatted template as a Stream

#### ToProvisioningTemplate(System.IO.Stream)
Method to convert a formatted template into a ProvisioningTemplate
> ##### Parameters
> **template:** The input formatted template as a Stream

> ##### Return value
> The output ProvisioningTemplate

#### ToProvisioningTemplate(System.IO.Stream,System.String)
Method to convert a formatted template into a ProvisioningTemplate, based on a specific ID
> ##### Parameters
> **template:** The input formatted template as a Stream

> **identifier:** The identifier of the template to convert

> ##### Return value
> The output ProvisioningTemplate

## Core.Framework.Provisioning.Providers.Json.JsonTemplateProvider
            
Provider for JSON based configurations
        

## Core.Framework.Provisioning.Providers.ITemplateProviderExtension
            
Interface for extending the XMLTemplateProvider while retrieving a template
        
### Properties

#### SupportsGetTemplatePreProcessing
Declares whether the object supports pre-processing during GetTemplate
#### SupportsGetTemplatePostProcessing
Declares whether the object supports post-processing during GetTemplate
#### SupportsSaveTemplatePreProcessing
Declares whether the object supports pre-processing during SaveTemplate
#### SupportsSaveTemplatePostProcessing
Declares whether the object supports post-processing during SaveTemplate
### Methods


#### Initialize(System.Object)
Initialization method to setup the extension object
> ##### Parameters
> **settings:** 


#### PreProcessGetTemplate(System.IO.Stream)
Method invoked before deserializing the template from the source repository
> ##### Parameters
> **stream:** The source stream

> ##### Return value
> The resulting stream, after pre-processing

#### PostProcessGetTemplate(OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate)
Method invoked after deserializing the template from the source repository
> ##### Parameters
> **template:** The just deserialized template

> ##### Return value
> The resulting template, after post-processing

#### PreProcessSaveTemplate(OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate)
Method invoked before serializing the template and before it is saved onto the target repository
> ##### Parameters
> **template:** The template that is going to be serialized

> ##### Return value
> The resulting template, after pre-processing

#### PostProcessSaveTemplate(System.IO.Stream)
Method invoked after serializing the template and before it is saved onto the target repository
> ##### Parameters
> **stream:** The source stream

> ##### Return value
> The resulting stream, after pre-processing

## Core.Framework.Provisioning.Connectors.AzureStorageConnector
            
Connector for files in Azure blob storage
        
### Methods


#### Constructor
Base constructor

#### Constructor
AzureStorageConnector constructor. Allows to directly set Azure Storage key and container
> ##### Parameters
> **connectionString:** Azure Storage Key (DefaultEndpointsProtocol=https;AccountName=yyyy;AccountKey=xxxx)

> **container:** Name of the Azure container to operate against


#### GetFiles
Get the files available in the default container
> ##### Return value
> List of files

#### GetFiles(System.String)
Get the files available in the specified container
> ##### Parameters
> **container:** Name of the container to get the files from

> ##### Return value
> List of files

#### GetFolders
Get the folders of the default container
> ##### Return value
> List of folders

#### GetFolders(System.String)
Get the folders of a specified container
> ##### Parameters
> **container:** Name of the container to get the folders from

> ##### Return value
> List of folders

#### GetFile(System.String)
Gets a file as string from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFile(System.String,System.String)
Gets a file as string from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String)
Gets a file as stream from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String,System.String)
Gets a file as stream from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### SaveFileStream(System.String,System.IO.Stream)
Saves a stream to the default container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **stream:** Stream containing the file contents


#### SaveFileStream(System.String,System.String,System.IO.Stream)
Saves a stream to the specified container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **container:** Name of the container to save the file to

> **stream:** Stream containing the file contents


#### DeleteFile(System.String)
Deletes a file from the default container
> ##### Parameters
> **fileName:** Name of the file to delete


#### DeleteFile(System.String,System.String)
Deletes a file from the specified container
> ##### Parameters
> **fileName:** Name of the file to delete

> **container:** Name of the container to delete the file from


## Core.Framework.Provisioning.Connectors.FileConnectorBase
            
Base file connector class
        
### Methods


#### GetFiles
Get the files available in the default container
> ##### Return value
> List of files

#### GetFiles(System.String)
Get the files available in the specified container
> ##### Parameters
> **container:** Name of the container to get the files from

> ##### Return value
> List of files

#### GetFolders
Get the folders of the default container
> ##### Return value
> List of folders

#### GetFolders(System.String)
Get the folders of a specified container
> ##### Parameters
> **container:** Name of the container to get the folders from

> ##### Return value
> List of folders

#### GetFile(System.String)
Gets a file as string from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFile(System.String,System.String)
Gets a file as string from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String)
Gets a file as stream from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String,System.String)
Gets a file as stream from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### SaveFileStream(System.String,System.IO.Stream)
Saves a stream to the default container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **stream:** Stream containing the file contents


#### SaveFileStream(System.String,System.String,System.IO.Stream)
Saves a stream to the specified container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **container:** Name of the container to save the file to

> **stream:** Stream containing the file contents


#### DeleteFile(System.String)
Deletes a file from the default container
> ##### Parameters
> **fileName:** Name of the file to delete


#### DeleteFile(System.String,System.String)
Deletes a file from the specified container
> ##### Parameters
> **fileName:** Name of the file to delete

> **container:** Name of the container to delete the file from


#### GetFilenamePart(System.String)
Returns a filename without a path
> ##### Parameters
> **fileName:** Path to the file to retrieve the filename from


## Core.Framework.Provisioning.Connectors.FileSystemConnector
            
Connector for files in file system
        
### Methods


#### Constructor
Base constructor

#### Constructor
FileSystemConnector constructor. Allows to directly set root folder and sub folder
> ##### Parameters
> **connectionString:** Root folder (e.g. c:\temp or .\resources or . or .\resources\templates)

> **container:** Sub folder (e.g. templates or resources\templates or blank


#### GetFiles
Get the files available in the default container
> ##### Return value
> List of files

#### GetFiles(System.String)
Get the files available in the specified container
> ##### Parameters
> **container:** Name of the container to get the files from

> ##### Return value
> List of files

#### GetFolders
Get the folders of the default container
> ##### Return value
> List of folders

#### GetFolders(System.String)
Get the folders of a specified container
> ##### Parameters
> **container:** Name of the container to get the folders from

> ##### Return value
> List of folders

#### GetFile(System.String)
Gets a file as string from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFile(System.String,System.String)
Gets a file as string from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String)
Gets a file as stream from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String,System.String)
Gets a file as stream from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### SaveFileStream(System.String,System.IO.Stream)
Saves a stream to the default container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **stream:** Stream containing the file contents


#### SaveFileStream(System.String,System.String,System.IO.Stream)
Saves a stream to the specified container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **container:** Name of the container to save the file to

> **stream:** Stream containing the file contents


#### DeleteFile(System.String)
Deletes a file from the default container
> ##### Parameters
> **fileName:** Name of the file to delete


#### DeleteFile(System.String,System.String)
Deletes a file from the specified container
> ##### Parameters
> **fileName:** Name of the file to delete

> **container:** Name of the container to delete the file from


## Core.Framework.Provisioning.Connectors.ICommitableFileConnector
            
Interface for File Connectors
        

## Core.Framework.Provisioning.Connectors.OpenXMLConnector
            
Connector that stores all the files into a unique .PNP OpenXML package
        
### Methods


#### Constructor
OpenXMLConnector constructor. Allows to manage a .PNP OpenXML package through an in memory stream.
> ##### Parameters
> **packageStream:** 


#### Constructor
OpenXMLConnector constructor. Allows to manage a .PNP OpenXML package file through a supporting persistence connector.
> ##### Parameters
> **packageFileName:** The name of the .PNP package file. If the .PNP extension is missing, it will be added

> **persistenceConnector:** The FileConnector object that will be used for physical persistence of the file

> **author:** The Author of the .PNP package file, if any. Optional

> **signingCertificate:** The X.509 certificate to use for digital signature of the template, optional


#### GetFiles
Get the files available in the default container
> ##### Return value
> List of files

#### GetFiles(System.String)
Get the files available in the specified container
> ##### Parameters
> **container:** Name of the container to get the files from (something like: "\images\subfolder")

> ##### Return value
> List of files

#### GetFolders
Get the folders of the default container
> ##### Return value
> List of folders

#### GetFolders(System.String)
Get the folders of a specified container
> ##### Parameters
> **container:** Name of the container to get the folders from

> ##### Return value
> List of folders

#### GetFile(System.String)
Gets a file as string from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFile(System.String,System.String)
Gets a file as string from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String)
Gets a file as stream from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String,System.String)
Gets a file as stream from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### SaveFileStream(System.String,System.IO.Stream)
Saves a stream to the default container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **stream:** Stream containing the file contents


#### SaveFileStream(System.String,System.String,System.IO.Stream)
Saves a stream to the specified container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **container:** Name of the container to save the file to

> **stream:** Stream containing the file contents


#### DeleteFile(System.String)
Deletes a file from the default container
> ##### Parameters
> **fileName:** Name of the file to delete


#### DeleteFile(System.String,System.String)
Deletes a file from the specified container
> ##### Parameters
> **fileName:** Name of the file to delete

> **container:** Name of the container to delete the file from


#### GetFileFromInsidePackage(System.String,System.String)
Will first try to find the file based on container/filename from the mapped file names. As a fallback it will try to find by container/filename in the pnp file structure, which was the original format.

## Core.Framework.Provisioning.Connectors.OpenXML.Model.PnPFileInfo
            
File descriptor for every single file in the PnP OpenXML file
        
### Properties

#### InternalName
The Internal Name of the file in the PnP OpenXML file
#### OriginalName
The Original Name of the file in the source template
#### Folder
The name of the folder within the PnP OpenXML file
#### Content
The binary content of the file

## Core.Framework.Provisioning.Connectors.OpenXML.Model.PnPFilesMap
            
Defines the mapping between original file names and OpenXML file names
        

## Core.Framework.Provisioning.Connectors.OpenXML.Model.PnPInfo
            
Global container of the PnP OpenXML file
        
### Properties

#### Manifest
The Manifest of the PnP OpenXML file
#### Properties
Custom properties of the PnP OpenXML file
#### Files
Files contained in the PnP OpenXML file
#### FilesMap
Defines the mapping between original file names and OpenXML file names

## Core.Framework.Provisioning.Connectors.OpenXML.Model.PnPManifest
            
Manifest of a PnP OpenXML file
        
### Properties

#### Type
The Type of the package file defined by the current manifest

## Core.Framework.Provisioning.Connectors.OpenXML.Model.PnPProperties
            
Properties of the PnP OpenXML container
        
### Properties

#### Id
Unique ID for the PnP OpenXML file
#### Author
Author of the PnP OpenXML file
#### CreationDateTime
Date and Time of creation for the PnP OpenXML file
#### Generator
Name of the Generator (engine) of the PnP OpenXML file

## Core.Framework.Provisioning.Connectors.OpenXML.PnPPackage
            
Defines a PnP OpenXML package file
        
### Properties

#### ManifestPart
The Manifest Part of the package file
#### Manifest
The Manifest of the package file
#### Properties
The Properties of the package
#### FilesMap
The File Map for files stored in the OpenXML file
#### FilesOriginPart
The Files origin
#### FilesPackageParts
The Files Parts of the package
#### Files
The Files of the package

## Core.Framework.Provisioning.Connectors.OpenXML.PnPPackageFileItem
            
Defines a single file in the PnP Open XML file package
        

## Core.Framework.Provisioning.Connectors.OpenXML.PnPPackageFormatException
            
Custom Exception type for PnP Packaging handling
        

## Core.Framework.Provisioning.Connectors.OpenXML.PnPPackageExtensions
            
Extension class for PnP OpenXML package files
        

## Core.Framework.Provisioning.Connectors.SharePointConnector
            
Connector for files in SharePoint
        
### Methods


#### Constructor
Base constructor

#### Constructor
SharePointConnector constructor. Allows to directly set root folder and sub folder
> ##### Parameters
> **clientContext:** 

> **connectionString:** Site collection URL (e.g. https://yourtenant.sharepoint.com/sites/dev)

> **container:** Library + folder that holds the files (mydocs/myfolder)


#### GetFiles
Get the files available in the default container
> ##### Return value
> List of files

#### GetFiles(System.String)
Get the files available in the specified container
> ##### Parameters
> **container:** Name of the container to get the files from

> ##### Return value
> List of files

#### GetFolders
Get the folders of the default container
> ##### Return value
> List of folders

#### GetFolders(System.String)
Get the folders of a specified container
> ##### Parameters
> **container:** Name of the container to get the folders from

> ##### Return value
> List of folders

#### GetFile(System.String)
Gets a file as string from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFile(System.String,System.String)
Gets a file as string from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String)
Gets a file as stream from the default container
> ##### Parameters
> **fileName:** Name of the file to get

> ##### Return value
> String containing the file contents

#### GetFileStream(System.String,System.String)
Gets a file as stream from the specified container
> ##### Parameters
> **fileName:** Name of the file to get

> **container:** Name of the container to get the file from

> ##### Return value
> String containing the file contents

#### SaveFileStream(System.String,System.IO.Stream)
Saves a stream to the default container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **stream:** Stream containing the file contents


#### SaveFileStream(System.String,System.String,System.IO.Stream)
Saves a stream to the specified container with the given name. If the file exists it will be overwritten
> ##### Parameters
> **fileName:** Name of the file to save

> **container:** Name of the container to save the file to

> **stream:** Stream containing the file contents


#### DeleteFile(System.String)
Deletes a file from the default container
> ##### Parameters
> **fileName:** Name of the file to delete


#### DeleteFile(System.String,System.String)
Deletes a file from the specified container
> ##### Parameters
> **fileName:** Name of the file to delete

> **container:** Name of the container to delete the file from


## Core.Framework.Provisioning.Extensibility.ExtensibilityManager
            
Provisioning Framework Component that is used for invoking custom providers during the provisioning process.
            
Provisioning Framework Component that is used for invoking custom providers during the provisioning process.
        
### Methods


#### ExecuteExtensibilityCallOut(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ExtensibilityHandler,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate)
Method to Invoke Custom Provisioning Providers. Ensure the ClientContext is not disposed in the custom provider.
> ##### Parameters
> **ctx:** Authenticated ClientContext that is passed to the custom provider.

> **handler:** A custom Extensibility Provisioning Provider

> **template:** ProvisioningTemplate that is passed to the custom provider

> ##### Exceptions
> **OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensiblityPipelineException:** 

> **System.ArgumentException:** Provider.Assembly or Provider.Type is NullOrWhiteSpace>

> **System.ArgumentNullException:** ClientContext is Null>


#### ExecuteTokenProviderCallOut(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ExtensibilityHandler,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate)
Method to Invoke Custom Provisioning Token Providers which implement the IProvisioningExtensibilityTokenProvider interface. Ensure the ClientContext is not disposed in the custom provider.
> ##### Parameters
> **ctx:** Authenticated ClientContext that is passed to the custom provider.

> **provider:** A custom Extensibility Provisioning Provider

> **template:** ProvisioningTemplate that is passed to the custom provider

> ##### Exceptions
> **OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensiblityPipelineException:** 

> **System.ArgumentException:** Provider.Assembly or Provider.Type is NullOrWhiteSpace>

> **System.ArgumentNullException:** ClientContext is Null>


#### ExecuteExtensibilityProvisionCallOut(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ExtensibilityHandler,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.ProvisioningTemplateApplyingInformation,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenParser,OfficeDevPnP.Core.Diagnostics.PnPMonitoredScope)
Ensure the ClientContext is not disposed in the custom provider.
Method to Invoke Custom Provisioning Handlers.
> ##### Parameters
> **ctx:** Authenticated ClientContext that is passed to the custom provider.

> **handler:** A custom Extensibility Provisioning Provider

> **template:** ProvisioningTemplate that is passed to the custom provider

> **applyingInformation:** The Provisioning Template application information object

> **tokenParser:** The Token Parser used by the engine during template provisioning

> **scope:** The PnPMonitoredScope of the current step in the pipeline

> ##### Exceptions
> **OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensiblityPipelineException:** 

> **System.ArgumentException:** Provider.Assembly or Provider.Type is NullOrWhiteSpace>

> **System.ArgumentNullException:** ClientContext is Null>


#### ExecuteExtensibilityExtractionCallOut(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ExtensibilityHandler,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.ProvisioningTemplateCreationInformation,OfficeDevPnP.Core.Diagnostics.PnPMonitoredScope)
Ensure the ClientContext is not disposed in the custom provider.
Method to Invoke Custom Extraction Handlers.
> ##### Parameters
> **ctx:** Authenticated ClientContext that is passed to the custom provider.

> **handler:** A custom Extensibility Provisioning Provider

> **template:** ProvisioningTemplate that is passed to the custom provider

> **creationInformation:** The Provisioning Template creation information object

> **scope:** The PnPMonitoredScope of the current step in the pipeline

> ##### Exceptions
> **OfficeDevPnP.Core.Framework.Provisioning.Extensibility.ExtensiblityPipelineException:** 

> **System.ArgumentException:** Provider.Assembly or Provider.Type is NullOrWhiteSpace>

> **System.ArgumentNullException:** ClientContext is Null>


## Core.Framework.Provisioning.Extensibility.IProvisioningExtensibilityProvider
            
Defines a interface that accepts requests from the provisioning processing component
        
### Methods


#### ProcessRequest(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,System.String)
Defines a interface that accepts requests from the provisioning processing component
> ##### Parameters
> **ctx:** 

> **template:** 

> **configurationData:** 


## Core.Framework.Provisioning.Extensibility.ExtensiblityPipelineException
            
Initializes a new instance of the ExtensiblityPipelineException class. This Exception occurs when there is an exception invoking a custom Extensibility Providers
        
### Methods


#### Constructor
Initializes a new instance of the ExtensiblityPipelineException class with a system supplied message

#### Constructor
Initializes a new instance of the ExtensiblityPipelineException class with the specified message string.
> ##### Parameters
> **message:** A string that describes the exception.


#### Constructor
Initializes a new instance of the ExtensiblityPipelineException class with a specified error message and a reference to the inner exception that is the cause of this exception.
> ##### Parameters
> **message:** A string that describes the exception.

> **innerException:** The exception that is the cause of the current exception.


#### Constructor
Initializes a new instance of the ExtensiblityPipelineException class from serialized data.
> ##### Parameters
> **info:** The object that contains the serialized data.

> **context:** The stream that contains the serialized data.

> ##### Exceptions
> **System.ArgumentNullException:** The info parameter is null.-or-The context parameter is null.


## Core.Framework.Provisioning.Extensibility.IProvisioningExtensibilityTokenProvider
            
Defines an interface which allows to plugin custom TokenDefinitions to the template provisioning pipeline
        
### Methods


#### GetTokens(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,System.String)
Provides Token Definitions to the template provisioning pipeline
> ##### Parameters
> **ctx:** 

> **template:** 

> **configurationData:** 


## Core.Framework.Provisioning.Extensibility.IProvisioningExtensibilityHandler
            
Defines an interface which allows to plugin custom Provisioning Extensibility Handlers to the template extraction/provisioning pipeline
        
### Methods


#### Provision(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.ProvisioningTemplateApplyingInformation,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenParser,OfficeDevPnP.Core.Diagnostics.PnPMonitoredScope,System.String)
Execute custom actions during provisioning of a template
> ##### Parameters
> **ctx:** The target ClientContext

> **template:** The current Provisioning Template

> **applyingInformation:** The Provisioning Template application information object

> **tokenParser:** Token parser instance

> **scope:** The PnPMonitoredScope of the current step in the pipeline

> **configurationData:** The configuration data, if any, for the handler


#### Extract(Microsoft.SharePoint.Client.ClientContext,OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate,OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.ProvisioningTemplateCreationInformation,OfficeDevPnP.Core.Diagnostics.PnPMonitoredScope,System.String)
Execute custom actions during extraction of a template
> ##### Parameters
> **ctx:** The target ClientContext

> **template:** The current Provisioning Template

> **creationInformation:** The Provisioning Template creation information object

> **scope:** The PnPMonitoredScope of the current step in the pipeline

> **configurationData:** The configuration data, if any, for the handler

> ##### Return value
> The Provisioning Template eventually enriched by the handler during extraction

## Core.Framework.Provisioning.Model.ComposedLook
            
Domain Object that defines a Composed Look in the Provision Template
        
### Properties

#### SiteLogo
Gets or sets the Site Logo
#### AlternateCSS
Gets or sets the AlternateCSS
#### MasterPage
Gets or sets the MasterPage for the Composed Look
#### Name
Gets or sets the Name
#### ColorFile
Gets or sets the ColorFile
#### FontFile
Gets or sets the FontFile
#### BackgroundFile
Gets or sets the Background Image
#### Version
Gets or sets the Version of the ComposedLook.

## Core.Framework.Provisioning.Model.ContentType
            
Domain Object used in the Provisioning template that defines a Content Type https://msdn.microsoft.com/en-us/library/office/ms463449.aspx
        
### Properties

#### Id
The Id of the Content Type
#### Name
The name of the Content Type
#### Description
The description of the Content Type
#### Group
The group name of the content type
#### FieldRefs
The FieldRefs entries of the List Instance
#### Hidden
True to define the content type as hidden. If you define a content type as hidden, SharePoint Foundation does not display that content type on the New button in list views.
#### Sealed
True to prevent changes to this content type. You cannot change the value of this attribute through the user interface, but you can change it in code if you have sufficient rights. You must have site collection administrator rights to unseal a content type.
#### ReadOnly
True to specify that the content type cannot be edited without explicitly removing the read-only setting. This can be done either in the user interface or in code.
#### Overwrite
True to overwrite an existing content type with the same ID.
#### DocumentTemplate
Specifies the document template for the content type
#### DocumentSetTemplate
Specifies the properties of the DocumentSet Template if the ContentType defines a DocumentSet
#### DisplayFormUrl
Specifies the URL of a custom display form to use for list items that have been assigned the content type
#### EditFormUrl
Specifies the URL of a custom edit form to use for list items that have been assigned the content type
#### NewFormUrl
Specifies the URL of a custom new form to use for list items that have been assigned the content type
#### 
Gets or Sets the Content Type ID
#### 
Gets or Sets if the Content Type should be the default Content Type in the library
#### 
Declares if the Content Type should be Removed from the list or library

## Core.Framework.Provisioning.Model.ContentTypeBinding
            
Domain Object for Content Type Binding in the Provisioning Template
        
### Properties

#### ContentTypeId
Gets or Sets the Content Type ID
#### Default
Gets or Sets if the Content Type should be the default Content Type in the library
#### Remove
Declares if the Content Type should be Removed from the list or library

## Core.Framework.Provisioning.Model.Feature
            
Domain Object that represents an Feature.
            
Domain Object that represents an Feature.
        
### Properties

#### Id
Gets or sets the feature Id
#### Deactivate
Gets or sets if the feature should be deactivated
#### 
A Collection of Features at the Site level
#### 
A Collection of Features at the Web level

## Core.Framework.Provisioning.Model.FieldRef
            
Represents a Field XML Markup that is used to define information about a field
            
Represents a Field XML Markup that is used to define information about a field
        
### Properties

#### Id
Gets ot sets the ID of the referenced field
#### Name
Gets or sets the name of the field link. This will not change the internal name of the field.
#### DisplayName
Gets or sets the Display Name of the field. Only applicable to fields associated with lists.
#### Required
Gets or sets if the field is Required
#### Hidden
Gets or sets if the field is Hidden

## Core.Framework.Provisioning.Model.ListInstance
            
This class holds deprecated ListInstance properties and methods
            
Domain Object that specifies the properties of the new list.
        
### Properties

#### Title
Gets or sets the list title
#### Description
Gets or sets the description of the list
#### DocumentTemplate
Gets or sets a value that specifies the identifier of the document template for the new list.
#### OnQuickLaunch
Gets or sets a value that specifies whether the new list is displayed on the Quick Launch of the site.
#### TemplateType
Gets or sets a value that specifies the list server template of the new list. https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.listtemplatetype.aspx
#### Url
Gets or sets a value that specifies whether the new list is displayed on the Quick Launch of the site.
#### EnableVersioning
Gets or sets whether verisioning is enabled on the list
#### EnableMinorVersions
Gets or sets whether minor verisioning is enabled on the list
#### DraftVersionVisibility
Gets or sets the DraftVersionVisibility for the list
#### EnableModeration
Gets or sets whether moderation/content approval is enabled on the list
#### MinorVersionLimit
Gets or sets the MinorVersionLimit for versioning, just in case it is enabled on the list
#### MaxVersionLimit
Gets or sets the MinorVersionLimit for verisioning, just in case it is enabled on the list
#### RemoveExistingContentTypes
Gets or sets whether existing content types should be removed
#### RemoveExistingViews
Gets or sets whether existing views should be removed
#### ContentTypesEnabled
Gets or sets whether content types are enabled
#### Hidden
Gets or sets whether to hide the list
#### ForceCheckout
Gets or sets whether to force checkout of documents in the library
#### EnableAttachments
Gets or sets whether attachments are enabled. Defaults to true.
#### EnableFolderCreation
Gets or sets whether folder is enabled. Defaults to true.
#### ContentTypeBindings
Gets or sets the content types to associate to the list
#### Views
Gets or sets the content types to associate to the list
#### FieldDefaults
Defines a list of default values for the Fields of the List Instance
#### Security
Defines the Security rules for the List Instance
#### Folders
Defines a collection of folders (eventually nested) that will be provisioned into the target list/library
#### UserCustomActions
Defines a collection of user custom actions that will be provisioned into the target list/library

## Core.Framework.Provisioning.Model.ProvisioningTemplate
            
Domain Object for the Provisioning Template
            
Domain Object for the Provisioning Template
        
### Properties

#### Providers
Gets a collection of Providers that are used during the extensibility pipeline
#### SearchSettings
The Search Settings for the Provisioning Template
#### Parameters
Any parameters that can be used throughout the template
#### Id
Gets or sets the ID of the Provisioning Template
#### Version
Gets or sets the Version of the Provisioning Template
#### SitePolicy
Gets or Sets the Site Policy
#### Security
Security Groups Members for the Template
#### Navigation
The Navigation configurations of the Provisioning Template
#### SiteFields
Gets a collection of fields
#### ContentTypes
Gets a collection of Content Types to create
#### Features
Gets or sets a list of features to activate or deactivate
#### CustomActions
Gets or sets CustomActions for the template
#### Files
Gets a collection of files for the template
#### Directories
Gets a collection of directories from which upload files for the template
#### ComposedLook
Gets or Sets the composed look of the template
#### Pages
Gets a collection of Wiki Pages for the template
#### TermGroups
Gets a collection of termgroups to deploy to the site
#### WebSettings
The Web Settings of the Provisioning Template
#### RegionalSettings
The Regional Settings of the Provisioning Template
#### SupportedUILanguages
The Supported UI Languages for the Provisioning Template
#### AuditSettings
The Audit Settings for the Provisioning Template
#### Workflows
Defines the Workflows to provision
#### SiteSearchSettings
The Site Collection level Search Settings for the Provisioning Template
#### WebSearchSettings
The Web level Search Settings for the Provisioning Template
#### AddIns
Defines the SharePoint Add-ins to provision
#### Publishing
Defines the Publishing configuration to provision
#### Properties
A set of custom Properties for the Provisioning Template
#### ImagePreviewUrl
The Image Preview Url of the Provisioning Template
#### DisplayName
The Display Name of the Provisioning Template
#### Description
The Description of the Provisioning Template
#### BaseSiteTemplate
The Base SiteTemplate of the Provisioning Template
#### 
References the parent ProvisioningTemplate for the current provisioning artifact
#### 
References the parent ProvisioningTemplate for the current provisioning artifact
### Methods


#### ToXML(OfficeDevPnP.Core.Framework.Provisioning.Providers.ITemplateFormatter)
Serializes a template to XML
> ##### Parameters
> **formatter:** 

> ##### Return value
> 

#### Constructor
Custom constructor to manage the ParentTemplate for the collection and all the children of the collection
> ##### Parameters
> **parentTemplate:** 


#### Constructor
Custom constructor to manage the ParentTemplate for the collection and all the children of the collection
> ##### Parameters
> **parentTemplate:** 


#### 
We implemented this to adhere to the generic List of T behavior
Finds an item matching a search predicate
> ##### Parameters
> **match:** The matching predicate to use for finding any target item

> ##### Return value
> The target item matching the find predicate

## Core.Framework.Provisioning.Model.BaseNavigationKind
            
Base abstract class for the navigation kinds (global or current)
        
### Properties

#### StructuralNavigation
Defines the Structural Navigation settings of the site
#### ManagedNavigation
Defines the Managed Navigation settings of the site

## Core.Framework.Provisioning.Model.CurrentNavigation
            
The Current Navigation settings for the Provisioning Template
        
### Fields

#### 
The site inherits the Global Navigation settings from its parent
#### 
The site uses Structural Global Navigation
#### 
The site uses Structural Local Current Navigation
#### 
The site uses Managed Global Navigation
### Properties

#### NavigationType
Defines the type of Current Navigation

## Core.Framework.Provisioning.Model.CurrentNavigationType
            
Defines the type of Current Navigation
        
### Fields

#### Inherit
The site inherits the Global Navigation settings from its parent
#### Structural
The site uses Structural Global Navigation
#### StructuralLocal
The site uses Structural Local Current Navigation
#### Managed
The site uses Managed Global Navigation

## Core.Framework.Provisioning.Model.CustomAction
            
Domain Object for custom actions associated with a SharePoint list, Web site, or subsite.
            
Domain Object for custom actions associated with a SharePoint list, Web site, or subsite.
        
### Properties

#### RightsValue
Gets or sets the value that specifies the permissions needed for the custom action. https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.permissionkind.aspx
#### Name
Gets or sets the name of the custom action.
#### Description
Gets or sets the description of the custom action.
#### Group
Gets or sets a value that specifies an implementation-specific value that determines the position of the custom action in the page.
#### Location
Gets or sets the location of the custom action. A string that contains the location; for example, Microsoft.SharePoint.SiteSettings.
#### Title
Gets or sets the display title of the custom action.
#### Sequence
Gets or sets the value that specifies an implementation-specific value that determines the order of the custom action that appears on the page.
#### Rights
Gets or sets the value that specifies the permissions needed for the custom action.
#### Url
Gets or sets the URL, URI, or ECMAScript (JScript, JavaScript) function associated with the action.
#### ScriptBlock
Gets or sets the value that specifies the ECMAScript to be executed when the custom action is performed.
#### ImageUrl
Gets or sets the URL of the image associated with the custom action.
#### ScriptSrc
Gets or sets a value that specifies the URI of a file which contains the ECMAScript to execute on the page
#### Remove
Gets or sets a value that specifies whether to Remove the CustomAction from the target
#### 
A Collection of CustomActions at the Site level
#### 
A Collection of CustomActions at the Web level

## Core.Framework.Provisioning.Model.Directory
            
Defines a Directory element, to describe a folder in the current repository that will be used to upload files into the target Site
        
### Properties

#### Src
The Src of the Directory
#### Folder
The TargetFolder of the Directory
#### Overwrite
The Overwrite flag for the files in the Directory
#### Level
The Level status for the files in the Directory
#### Recursive
Defines whether to recursively browse through all the child folders of the Directory
#### IncludedExtensions
The file Extensions to include while uploading the Directory
#### ExcludedExtensions
The file Extensions to exclude while uploading the Directory
#### MetadataMappingFile
The file path of JSON mapping file with metadata for files to upload in the Directory
#### Security
Defines the Security rules for the File

## Core.Framework.Provisioning.Model.DirectoryCollection
            
Collection of Directory objects
        

## Core.Framework.Provisioning.Model.ExtensibilityHandler
            
Domain Object for Extensiblity Call out
        

## Core.Framework.Provisioning.Model.FeatureCollection
            
Collection of Feature objects
        

## Core.Framework.Provisioning.Model.AddIn
            
Defines an Add-in to provision
        
### Properties

#### PackagePath
Defines the .app file of the SharePoint Add-in to provision
#### Source
Defines the Source of the SharePoint Add-in to provision Possible values are: CorporateCatalog, DeveloperSite, InvalidSource, Marketplace, ObjectModel, RemoteObjectModel

## Core.Framework.Provisioning.Model.AddInCollection
            
A collection of AddIn objects
        

## Core.Framework.Provisioning.Model.AuditSettings
            
The Audit Settings for the Provisioning Template
        
### Properties

#### AuditFlags
Audit Flags configured for the Site
#### AuditLogTrimmingRetention
The Audit Log Trimming Retention for Audits
#### TrimAuditLog
A flag to enable Audit Log Trimming

## Core.Framework.Provisioning.Model.AvailableWebTemplate
            
Defines an available Web Template for the current Publishing site
        
### Properties

#### LanguageCode
The Language Code for the Web Template
#### TemplateName
The Name of the Web Template

## Core.Framework.Provisioning.Model.AvailableWebTemplateCollection
            
A collection of AvailableWebTemplate objects
        

## Core.Framework.Provisioning.Model.BaseModel
            
Base type for any Domain Model object (excluded the ProvisioningTemplate type)
        
### Properties

#### ParentTemplate
References the parent ProvisioningTemplate for the current provisioning artifact

## Core.Framework.Provisioning.Model.ContentTypeBindingCollection
            
Collection of ContentTypeBinding objects
        

## Core.Framework.Provisioning.Model.ContentTypeCollection
            
Collection of ContentType objects
        

## Core.Framework.Provisioning.Model.CustomActionCollection
            
Collection of CustomAction objects
        

## Core.Framework.Provisioning.Model.CustomActions
            
Domain Object that represents a Collections of Custom Actions
        
### Properties

#### SiteCustomActions
A Collection of CustomActions at the Site level
#### WebCustomActions
A Collection of CustomActions at the Web level

## Core.Framework.Provisioning.Model.DataRowCollection
            
Collection of DataRow objects
        

## Core.Framework.Provisioning.Model.DefaultDocument
            
A default document for a Document Set
        
### Properties

#### Name
The name (including the relative path) of the Default Document for a Document Set
#### ContentTypeId
The value of the ContentTypeID of the Default Document for the Document Set
#### FileSourcePath
The path of the file to upload as a Default Document for the Document Set

## Core.Framework.Provisioning.Model.DesignPackage
            
Defines a Design Package to import into the current Publishing site
        
### Properties

#### DesignPackagePath
Defines the path of the Design Package to import into the current Publishing site
#### MajorVersion
The Major Version of the Design Package to import into the current Publishing site
#### MinorVersion
The Minor Version of the Design Package to import into the current Publishing site
#### PackageGuid
The ID of the Design Package to import into the current Publishing site
#### PackageName
The Name of the Design Package to import into the current Publishing site

## Core.Framework.Provisioning.Model.DocumentSetTemplate
            
Defines a DocumentSet Template for creating multiple DocumentSet instances
        
### Properties

#### AllowedContentTypes
The list of allowed Content Types for the Document Set
#### DefaultDocuments
The list of default Documents for the Document Set
#### SharedFields
The list of Shared Fields for the Document Set
#### WelcomePageFields
The list of Welcome Page Fields for the Document Set
#### WelcomePage
Defines the custom WelcomePage for the Document Set

## Core.Framework.Provisioning.Model.Features
            
Domain Object that is used in the Site Template for OOB Features
        
### Properties

#### SiteFeatures
A Collection of Features at the Site level
#### WebFeatures
A Collection of Features at the Web level

## Core.Framework.Provisioning.Model.Field
            
Represents a Field XML Markup that is used to define information about a field
        
### Properties

#### 
Gets ot sets the ID of the referenced field
#### 
Gets or sets the name of the field link. This will not change the internal name of the field.
#### 
Gets or sets the Display Name of the field. Only applicable to fields associated with lists.
#### 
Gets or sets if the field is Required
#### 
Gets or sets if the field is Hidden
#### SchemaXml
Gets a value that specifies the XML Schema representing the Field type. https://msdn.microsoft.com/en-us/library/office/ff407271.aspx

## Core.Framework.Provisioning.Model.FieldRefCollection
            
Collection of FieldRef objects
        

## Core.Framework.Provisioning.Model.FieldCollection
            
Collection of Field objects
        

## Core.Framework.Provisioning.Model.FileCollection
            
Collection of File objects
        

## Core.Framework.Provisioning.Model.FileLevel
            
The File Level for a File element
        
### Fields

#### Draft
The file will be stored as a draft
#### Checkout
The file will be stored as a checked out item
#### Published
The file will be stored as a published item

## Core.Framework.Provisioning.Model.Folder
            
Defines a folder that will be provisioned into the target list/library
        
### Properties

#### Name
The Name of the Folder
#### Security
Defines the security rules for the current Folder
#### Folders
Defines the child folders of the current Folder, if any

## Core.Framework.Provisioning.Model.FolderCollection
            
Collection of Folder objects
        

## Core.Framework.Provisioning.Model.GlobalNavigation
            
The Global Navigation settings for the Provisioning Template
        
### Fields

#### 
The site inherits the Global Navigation settings from its parent
#### 
The site uses Structural Global Navigation
#### 
The site uses Managed Global Navigation
### Properties

#### NavigationType
Defines the type of Global Navigation

## Core.Framework.Provisioning.Model.GlobalNavigationType
            
Defines the type of Global Navigation
        
### Fields

#### Inherit
The site inherits the Global Navigation settings from its parent
#### Structural
The site uses Structural Global Navigation
#### Managed
The site uses Managed Global Navigation

## Core.Framework.Provisioning.Model.IProvisioningTemplateDescendant
            
Interface implemented by any descendant of a ProvisioningTemplate
        
### Properties

#### ParentTemplate
References the parent ProvisioningTemplate for the current provisioning artifact

## Core.Framework.Provisioning.Model.ListInstanceCollection
            
Collection of ListInstance objects
        

## Core.Framework.Provisioning.Model.Localization
            
Domain Object used in the Provisioning template that defines a Localization item
        
### Properties

#### LCID
The Locale ID of a Localization Language
#### Name
The Name of a Localization Language
#### ResourceFile
The path to the .RESX (XML) resource file for the current Localization

## Core.Framework.Provisioning.Model.LocalizationCollection
            
Collection of Localization objects
        

## Core.Framework.Provisioning.Model.ManagedNavigation
            
Defines the Managed Navigation settings of a site
        
### Properties

#### TermStoreId
Defines the TermStore ID for the Managed Navigation
#### TermSetId
Defines the TermSet ID for the Managed Navigation

## Core.Framework.Provisioning.Model.Navigation
            
The Navigation configurations of the Provisioning Template
        
### Properties

#### GlobalNavigation
The Global Navigation settings for the Provisioning Template
#### CurrentNavigation
The Current Navigation settings for the Provisioning Template
#### 
A collection of navigation nodes children of the current NavigatioNode
#### 
Defines the Title of a Navigation Node
#### 
Defines the Url of a Navigation Node
#### 
Defines whether the Navigation Node for the Structural Navigation targets an External resource

## Core.Framework.Provisioning.Model.NavigationNode
            
Defines a Navigation Node for the Structural Navigation of a site
        
### Properties

#### NavigationNodes
A collection of navigation nodes children of the current NavigatioNode
#### Title
Defines the Title of a Navigation Node
#### Url
Defines the Url of a Navigation Node
#### IsExternal
Defines whether the Navigation Node for the Structural Navigation targets an External resource

## Core.Framework.Provisioning.Model.NavigationNodeCollection
            
A collection of NavigationNode objects
        

## Core.Framework.Provisioning.Model.PageLayout
            
Defines an available Page Layout for the current Publishing site
        
### Properties

#### Path
Defines the path of the Page Layout for the current Publishing site
#### IsDefault
Defines whether the Page Layout is the default for the current Publishing site

## Core.Framework.Provisioning.Model.PageLayoutCollection
            
Collection of PageLayout objects
        

## Core.Framework.Provisioning.Model.PageCollection
            
Collection of Page objects
        

## Core.Framework.Provisioning.Model.PropertyBagEntryCollection
            
Collection of PropertyBagEntry objects
        

## Core.Framework.Provisioning.Model.ExtensibilityHandlerCollection
            
Collection of ExtensibilityHandler objects
        

## Core.Framework.Provisioning.Model.ProviderCollection
            
Collection of Provider objects
        

## Core.Framework.Provisioning.Model.ProvisioningTemplateDictionary`2
            
Generic keyed collection of items stored in the ProvisioningTemplate graph
            The type of the Key for the keyed collection
            The type of the Item for the keyed collection
        
### Properties

#### ParentTemplate
References the parent ProvisioningTemplate for the current provisioning artifact
### Methods


#### Constructor
Custom constructor to manage the ParentTemplate for the collection and all the children of the collection
> ##### Parameters
> **parentTemplate:** 


## Core.Framework.Provisioning.Model.ProvisioningTemplateCollection`1
            
Generic collection of items stored in the ProvisioningTemplate graph
            The type of Item for the collection
        
### Properties

#### ParentTemplate
References the parent ProvisioningTemplate for the current provisioning artifact
### Methods


#### Constructor
Custom constructor to manage the ParentTemplate for the collection and all the children of the collection
> ##### Parameters
> **parentTemplate:** 


#### Find(System.Predicate{`0})
We implemented this to adhere to the generic List of T behavior
Finds an item matching a search predicate
> ##### Parameters
> **match:** The matching predicate to use for finding any target item

> ##### Return value
> The target item matching the find predicate

## Core.Framework.Provisioning.Model.Publishing
            
Defines the Publishing configuration to provision
        
### Properties

#### DesignPackage
Defines a Design Package to import into the current Publishing site
#### AvailableWebTemplates
Defines the Available Web Templates for the current Publishing site
#### PageLayouts
Defines the Available Page Layouts for the current Publishing site
#### AutoCheckRequirements
Defines how an engine should behave if the requirements for provisioning publishing capabilities are not satisfied by the target site

## Core.Framework.Provisioning.Model.AutoCheckRequirementsOptions
            
Defines how an engine should behave if the requirements for provisioning publishing capabilities are not satisfied by the target site
        
### Fields

#### MakeCompliant
Instructs the engine to make the target site compliant with the requirements
#### SkipIfNotCompliant
Instructs the engine to skip the Publishing section if the target site is not compliant with the requirements
#### FailIfNotCompliant
Instructs the engine to throw an exception/failure if the target site is not compliant with the requirements

## Core.Framework.Provisioning.Model.RoleAssignment
            
Role Assignment for a target Principal
        
### Properties

#### Principal
Defines the Role to which the assignment will apply
#### RoleDefinition
Defines the Role to which the assignment will apply

## Core.Framework.Provisioning.Model.RoleAssignmentCollection
            
Collection of RoleAssignment objects
        

## Core.Framework.Provisioning.Model.RoleDefinitionCollection
            
Collection of RoleDefinition objects
        

## Core.Framework.Provisioning.Model.SiteGroupCollection
            
Collection of SiteGroup objects
        

## Core.Framework.Provisioning.Model.SiteSecurityPermissions
            
Permission settings for the target Site
        
### Properties

#### RoleDefinitions
List of Role Definitions for the Site
#### RoleAssignments
List of Role Assignments for the Site

## Core.Framework.Provisioning.Model.RegionalSettings
            
Defines the Regional Settings for a site
        
### Properties

#### AdjustHijriDays
The number of days to extend or reduce the current month in Hijri calendars
#### AlternateCalendarType
The Alternate Calendar type that is used on the server
#### CalendarType
The Calendar Type that is used on the server
#### Collation
The Collation that is used on the site
#### FirstDayOfWeek
The First Day of the Week used in calendars on the server
#### FirstWeekOfYear
The First Week of the Year used in calendars on the server
#### LocaleId
The Locale Identifier in use on the server
#### ShowWeeks
Defines whether to display the week number in day or week views of a calendar
#### Time24
Defines whether to use a 24-hour time format in representing the hours of the day
#### TimeZone
The Time Zone that is used on the server
#### WorkDayEndHour
The the default hour at which the work day ends on the calendar that is in use on the server
#### WorkDays
The work days of Web site calendars
#### WorkDayStartHour
The the default hour at which the work day starts on the calendar that is in use on the server

## Core.Framework.Provisioning.Model.WorkHour
            
The Work Hours of a Day
        

## Core.Framework.Provisioning.Model.SiteGroup
            
The base type for a Site Group
        
### Properties

#### Members
The list of members of the Site Group
#### Title
The Title of the Site Group
#### Description
The Description of the Site Group
#### Owner
The Owner of the Site Group
#### AllowMembersEditMembership
Defines whether the members can edit membership of the Site Group
#### AllowRequestToJoinLeave
Defines whether to allow requests to join or leave the Site Group
#### AutoAcceptRequestToJoinLeave
Defines whether to auto-accept requests to join or leave the Site Group
#### OnlyAllowMembersViewMembership
Defines whether to allow members only to view the membership of the Site Group
#### RequestToJoinLeaveEmailSetting
Defines the email address used for membership requests to join or leave will be sent for the Site Group

## Core.Framework.Provisioning.Model.StructuralNavigation
            
Defines the Structural Navigation settings of a site
        
### Properties

#### RemoveExistingNodes
Defines whether to remove existing nodes before creating those described through this element
#### NavigationNodes
A collection of navigation nodes for the site

## Core.Framework.Provisioning.Model.SupportedUILanguage
            
Defines a single Supported UI Language for a site
        
### Properties

#### LCID
The Locale ID of a Supported UI Language

## Core.Framework.Provisioning.Model.SupportedUILanguageCollection
            
Collection of SupportedUILanguage objects
        

## Core.Framework.Provisioning.Model.TermCollection
            
Collection of Term objects
        

## Core.Framework.Provisioning.Model.TermGroupCollection
            
Collection of TermGroup objects
        

## Core.Framework.Provisioning.Model.TermLabelCollection
            
Collection of TermLabel objects
        

## Core.Framework.Provisioning.Model.TermSetCollection
            
Collection of TermSete objects
        

## Core.Framework.Provisioning.Model.UserCollection
            
Collection of User objects
        

## Core.Framework.Provisioning.Model.ViewCollection
            
Collection of View objects
        

## Core.Framework.Provisioning.Model.File
            
Defines a File element, to describe a file that will be provisioned into the target Site
        
### Fields

#### 
The file will be stored as a draft
#### 
The file will be stored as a checked out item
#### 
The file will be stored as a published item
### Properties

#### Src
The Src of the File
#### Folder
The TargetFolder of the File
#### Overwrite
The Overwrite flag for the File
#### Level
The Level status for the File
#### Security
Defines the Security rules for the File

## Core.Framework.Provisioning.Model.Provider
            
Domain Object for Extensiblity Call out
        

## Core.Framework.Provisioning.Model.SiteSecurity
            
Domain Object that is used in the site template
        
### Properties

#### 
List of Role Definitions for the Site
#### 
List of Role Assignments for the Site
#### AdditionalAdministrators
A Collection of users that are associated as site collection adminsitrators
#### AdditionalOwners
A Collection of users that are associated to the sites owners group
#### AdditionalMembers
A Collection of users that are associated to the sites members group
#### AdditionalVisitors
A Collection of users taht are associated to the sites visitors group
#### SiteGroups
List of additional Groups for the Site
#### SiteSecurityPermissions
List of Site Security Permissions for the Site
#### BreakRoleInheritance
Declares whether the to break role inheritance for the site, if it is a sub-site
#### CopyRoleAssignments
Defines whether to copy role assignments or not while breaking role inheritance
#### ClearSubscopes
Defines whether to clear subscopes or not while breaking role inheritance for the site

## Core.Framework.Provisioning.Model.User
            
Domain Object that defines a User or group in the provisioning template
        
### Properties

#### Name
The User email Address or the group name.

## Core.Framework.Provisioning.Model.WebPartCollection
            
Collection of WebPart objects
        

## Core.Framework.Provisioning.Model.WebSettings
            
Domain Object used in the Provisioning template that defines a Section of Settings for the current Web Site
        
### Properties

#### NoCrawl
Defines whether the site has to be crawled or not
#### RequestAccessEmail
The email address to which any access request will be sent
#### WelcomePage
Defines the Welcome Page (Home Page) of the site to which the Provisioning Template is applied.
#### Title
The Title of the Site, optional attribute.
#### Description
The Description of the Site, optional attribute.
#### SiteLogo
The SiteLogo of the Site, optional attribute.
#### AlternateCSS
The AlternateCSS of the Site, optional attribute.
#### MasterPageUrl
The MasterPage Url of the Site, optional attribute.
#### CustomMasterPageUrl
The Custom MasterPage Url of the Site, optional attribute.

## Core.Framework.Provisioning.Model.WorkflowDefinition
            
Defines a Workflow Definition to provision
        
### Properties

#### Properties
Defines the Properties of the Workflows to provision
#### FormField
Defines the FormField XML of the Workflow to provision
#### Id
Defines the ID of the Workflow Definition for the current Subscription
#### AssociationUrl
Defines the URL of the Workflow Association page
#### Description
The Description of the Workflow
#### DisplayName
The Display Name of the Workflow
#### DraftVersion
Defines the DraftVersion of the Workflow, optional attribute.
#### InitiationUrl
Defines the URL of the Workflow Initiation page
#### Published
Defines if the Workflow is Published, optional attribute.
#### RequiresAssociationForm
Defines if the Workflow requires the Association Form
#### RequiresInitiationForm
Defines if the Workflow requires the Initiation Form
#### RestrictToScope
Defines the Scope Restriction for the Workflow
#### RestrictToType
Defines the Type of Scope Restriction for the Workflow
#### XamlPath
Defines path of the XAML of the Workflow to provision

## Core.Framework.Provisioning.Model.WorkflowDefinitionCollection
            
Defines a collection of objects of type WorkflowDefinition
        

## Core.Framework.Provisioning.Model.Workflows
            
Defines the Workflows to provision
        
### Properties

#### WorkflowDefinitions
Defines the Workflows Definitions to provision
#### WorkflowSubscriptions
Defines the Workflows Subscriptions to provision

## Core.Framework.Provisioning.Model.WorkflowSubscription
            
Defines a Workflow Subscription to provision
        
### Properties

#### PropertyDefinitions
Defines the Property Definitions of the Workflows to provision
#### DefinitionId
Defines the ID of the Workflow Definition for the current Subscription
#### ListId
Defines the ID of the target list/library for the current Subscription, Optional and if it is missing, the workflow subscription will be at Site level
#### Enabled
Defines if the Workflow Definition is enabled for the current Subscription
#### EventSourceId
Defines the ID of the Event Source for the current Subscription
#### EventTypes
Defines the list of events that will start the workflow instance Possible values in the list: WorkflowStartEvent, ItemAddedEvent, ItemUpdatedEvent
#### ManualStartBypassesActivationLimit
Defines if the Workflow can be manually started bypassing the activation limit
#### Name
Defines the Name of the Workflow Subscription
#### ParentContentTypeId
Defines the Parent ContentType Id of the Workflow Subscription
#### StatusFieldName
Defines the Status Field Name of the Workflow Subscription

## Core.Framework.Provisioning.Model.WorkflowSubscriptionCollection
            
Defines a collection of objects of type WorkflowSubscription
        

## Core.Framework.Provisioning.ObjectHandlers.ObjectExtensibilityHandlers
            
Extensibility Provider CallOut
        

## Core.Framework.TimerJobs.Enums.AuthenticationType
            
Type of authentication, supports Office365, NetworkCredentials (on-premises) and AppOnly (both Office 365 as On-premises)
        

## Core.Framework.TimerJobs.TimerJobRunHandler
            
TimerJobRun delegate
            
> **sender:** calling object instance

            
> **e:** TimerJobRunEventArgs event arguments instance

        

## Core.Framework.TimerJobs.TimerJob
            
Abstract base class for creating timer jobs (background processes) that operate against SharePoint sites. These timer jobs are designed to use the CSOM API and thus can run on any server that can communicate with SharePoint.
        
### Fields

#### 
DateTime of the previous run attempt
#### 
Bool indicating if the previous run was successful
#### 
Timer job version used during the previous run
#### 
Property value collection used to store timer job custom properties
### Properties

#### Name
Gets the name of this timer job
#### Version
Gets the version of this timer job
#### ConfigurationData
Gets or sets additional timer job configuration data
#### ManageState
Gets and sets the state management value: when true the timer job will automatically handle state by storing a json serialized class as a web property bag entry. Default value is false
#### IsRunning
Is this timer job running?
#### UseThreading
Can this timer job use multiple threads. Defaults to true
#### MaximumThreads
How many threads can be used by this timer job. Default value is 5.
#### AuthenticationType
Gets the authentication type that the timer job will use. This will be set as part of the UseOffice365Authentication and UseNetworkCredentialsAuthentication methods
#### SharePointVersion
Gets or sets the SharePoint version. Default value is detected based on the laoded CSOM assembly version, but can be overriden in case you want to for example use v16 assemblies in v15 (on-premises)
#### Realm
Realm will be automatically defined, but there's an option to manually specify it which may be needed when did an override of ResolveAddedSites and specify your sites.
#### TenantAdminSite
Option to specify the tenant admin site. For MT this typically is not needed since we can detect the tenant admin site, but for on premises and DvNext this is needed
#### ExpandSubSites
Does the timerjob need to fire as well for every sub site in the site?
#### EnumerationUser
Returns the user account used for enumaration. Enumeration is done using search and the search API requires a user context
#### EnumerationPassword
Returns the password of the user account used for enumaration. Enumeration is done using search and the search API requires a user context
#### EnumerationDomain
Returns the domain of the user account used for enumaration. Enumeration is done using search and the search API requires a user context
### Methods


#### Constructor
Simpliefied constructor for timer job, version is always set to "1.0"
> ##### Parameters
> **name:** Name of the timer job


#### Constructor
Default constructor for timer job
> ##### Parameters
> **name:** Name of the timer job

> **version:** Version of the timer job

> **configurationData:** 


#### Run
Triggers the timer job to start running

#### DoWorkBatch(System.Collections.Generic.List{System.String})
Processes the amount of work that will be done by one thread
> ##### Parameters
> **sites:** Batch of sites that the thread will need to process


#### DoWork(System.String)
Processes the amount of work that will be done for a single site/web
> ##### Parameters
> **site:** Url of the site to process


#### OnTimerJobRun(OfficeDevPnP.Core.Framework.TimerJobs.TimerJobRunEventArgs)
Triggers the event to fire and deals with all the pre/post processing needed to automatically manage state
> ##### Parameters
> **e:** TimerJobRunEventArgs event arguments class that will be passed to the event handler


#### CreateWorkBatches
Creates batches of sites to process. Batch size is based on max number of threads
> ##### Return value
> List of Lists holding the work batches

#### UseOffice365Authentication(System.String,System.String)
Prepares the timerjob to operate against Office 365 with user and password credentials. Sets AuthenticationType to AuthenticationType.Office365
> ##### Parameters
> **userUPN:** 

> **password:** Password of the user that will be used to operate the timer job work


#### UseOffice365Authentication(System.String,System.Security.SecureString)
Prepares the timerjob to operate against Office 365 with user and password credentials. Sets AuthenticationType to AuthenticationType.Office365
> ##### Parameters
> **userUPN:** 

> **password:** Password of the user that will be used to operate the timer job work


#### UseOffice365Authentication(System.String)
Prepares the timerjob to operate against Office 365 with user and password credentials which are retrieved via the windows Credential Manager. Also sets AuthenticationType to AuthenticationType.Office365
> ##### Parameters
> **credentialName:** Name of the credential manager registration


#### UseNetworkCredentialsAuthentication(System.String,System.String,System.String)
Prepares the timerjob to operate against SharePoint on-premises with user name password credentials. Sets AuthenticationType to AuthenticationType.NetworkCredentials
> ##### Parameters
> **samAccountName:** samAccontName of the windows user

> **password:** Password of the windows user

> **domain:** NT domain of the windows user


#### UseNetworkCredentialsAuthentication(System.String,System.Security.SecureString,System.String)
Prepares the timerjob to operate against SharePoint on-premises with user name password credentials. Sets AuthenticationType to AuthenticationType.NetworkCredentials
> ##### Parameters
> **samAccountName:** samAccontName of the windows user

> **password:** Password of the windows user

> **domain:** NT domain of the windows user


#### UseNetworkCredentialsAuthentication(System.String)
Prepares the timerjob to operate against SharePoint on-premises with user name password credentials which are retrieved via the windows Credential Manager. Sets AuthenticationType to AuthenticationType.NetworkCredentials
> ##### Parameters
> **credentialName:** Name of the credential manager registration


#### UseAppOnlyAuthentication(System.String,System.String)
Prepares the timerjob to operate against SharePoint on-premises with app-only credentials. Sets AuthenticationType to AuthenticationType.AppOnly
> ##### Parameters
> **clientId:** Client ID of the app

> **clientSecret:** Client Secret of the app


#### UseAzureADAppOnlyAuthentication(System.String,System.String,System.String,System.String)
Prepares the timerjob to operate against SharePoint Only with Azure AD app-only credentials. Sets AuthenticationType to AuthenticationType.AzureADAppOnly
> ##### Parameters
> **clientId:** Client ID of the app

> **azureTenant:** The Azure tenant name, like contoso.com

> **certificatePath:** The path to the *.pfx certicate file

> **certificatePassword:** The password to the certificate


#### UseAzureADAppOnlyAuthentication(System.String,System.String,System.String,System.Security.SecureString)
Prepares the timerjob to operate against SharePoint Only with Azure AD app-only credentials. Sets AuthenticationType to AuthenticationType.AzureADAppOnly
> ##### Parameters
> **clientId:** Client ID of the app

> **azureTenant:** The Azure tenant name, like contoso.com

> **certificatePath:** The path to the *.pfx certicate file

> **certificatePassword:** The password to the certificate


#### UseAzureADAppOnlyAuthentication(System.String,System.String,System.Security.Cryptography.X509Certificates.X509Certificate2)
Prepares the timerjob to operate against SharePoint Only with Azure AD app-only credentials. Sets AuthenticationType to AuthenticationType.AzureADAppOnly
> ##### Parameters
> **clientId:** Client ID of the app

> **azureTenant:** The Azure tenant name, like contoso.com

> **certificate:** The X.509 Certificate to use for AppOnly Authentication


#### Clone(OfficeDevPnP.Core.Framework.TimerJobs.TimerJob)
Takes over the settings from the passed timer job. Is useful when you run multiple jobs in a row or chain job execution. Settings that are taken over are all the authentication, enumeration settings and SharePointVersion
> ##### Parameters
> **job:** 


#### GetAuthenticationManager(System.String)
Get an AuthenticationManager instance per host Url. Needed to make this work properly, else we're getting access denied because of Invalid audience Uri
> ##### Parameters
> **url:** Url of the site

> ##### Return value
> An instantiated AuthenticationManager

#### SetEnumerationCredentials(System.String,System.String)
Provides the timer job with the enumeration credentials. For Office 365 username and password is sufficient
> ##### Parameters
> **userUPN:** 

> **password:** Password of the enumeration user


#### SetEnumerationCredentials(System.String,System.Security.SecureString)
Provides the timer job with the enumeration credentials. For Office 365 username and password is sufficient
> ##### Parameters
> **userUPN:** 

> **password:** Password of the enumeration user


#### SetEnumerationCredentials(System.String,System.String,System.String)
Provides the timer job with the enumeration credentials. For SharePoint on-premises username, password and domain are needed
> ##### Parameters
> **samAccountName:** UPN of the enumeration user

> **password:** Password of the enumeration user

> **domain:** Domain of the enumeration user


#### SetEnumerationCredentials(System.String,System.Security.SecureString,System.String)
Provides the timer job with the enumeration credentials. For SharePoint on-premises username, password and domain are needed
> ##### Parameters
> **samAccountName:** Account name of the enumeration user

> **password:** Password of the enumeration user

> **domain:** Domain of the enumeration user


#### SetEnumerationCredentials(System.String)
Provides the timer job with the enumeration credentials. For SharePoint on-premises username, password and domain are needed
> ##### Parameters
> **credentialName:** Name of the credential manager registration


#### AddSite(System.String)
Adds a site Url or wildcard site Url to the collection of sites that the timer job will process
> ##### Parameters
> **site:** Site Url or wildcard site Url to be processed by the timer job


#### ClearAddedSites
Clears the list of added site Url's and/or wildcard site Url's

#### UpdateAddedSites(System.Collections.Generic.List{System.String})
Virtual method that can be overriden to allow the timer job itself to control the list of sites to operate against. Scenario is for example timer job that reads this data from a database instead of being fed by the calling program
> ##### Parameters
> **addedSites:** List of added site Url's and/or wildcard site Url's

> ##### Return value
> List of added site Url's and/or wildcard site Url's

#### ResolveAddedSites(System.Collections.Generic.List{System.String})
Virtual method that can be overriden to control the list of resolved sites
> ##### Parameters
> **addedSites:** List of added site Url's and/or wildcard site Url's

> ##### Return value
> List of resolved sites

#### DoExpandBatch(System.Collections.Generic.List{System.String},System.Collections.Generic.List{System.String})
Processes one bach of sites to expand, whcih is the workload of one thread
> ##### Parameters
> **sites:** Batch of sites to expand

> **resolvedSitesAndSubSites:** List holding the expanded sites


#### CreateExpandBatches(System.Collections.Generic.List{System.String})
Creates batches of sites to expand
> ##### Parameters
> **resolvedSites:** List of sites to expand

> ##### Return value
> List of list with batches of sites to expand

#### ExpandSite(System.Collections.Generic.List{System.String},System.String)
Expands and individual site into sub sites
> ##### Parameters
> **resolvedSitesAndSubSites:** list of sites and subsites resulting from the expanding

> **site:** site to expand


#### CreateClientContext(System.String)
Creates a ClientContext object based on the set AuthenticationType and the used version of SharePoint
> ##### Parameters
> **site:** Site Url to create a ClientContext for

> ##### Return value
> The created ClientContext object. Returns null if no ClientContext was created

#### ResolveSite(System.String,System.Collections.Generic.List{System.String})
Resolves a wildcard site Url into a list of actual site Url's
> ##### Parameters
> **site:** Wildcard site Url to resolve

> **resolvedSites:** List of resolved site Url's


#### GetAllSubSites(Microsoft.SharePoint.Client.Site)
Gets all sub sites for a given site
> ##### Parameters
> **site:** Site to find all sub site for

> ##### Return value
> IEnumerable of strings holding the sub site urls

#### IsValidUrl(System.String)
Verifies if the passed Url has a valid structure
> ##### Parameters
> **url:** Url to validate

> ##### Return value
> True is valid, false otherwise

#### GetSharePointVersion
Gets the current SharePoint version based on the loaded assembly
> ##### Return value
> 

#### GetTenantAdminSite(System.String)
Gets the tenant admin site based on the tenant name provided when setting the authentication details
> ##### Return value
> The tenant admin site

#### GetTopLevelSite(System.String)
Gets the top level site for the given url
> ##### Parameters
> **site:** 

> ##### Return value
> 

#### GetRootSite(System.String)
Gets the root site for a given site Url
> ##### Parameters
> **site:** Site Url

> ##### Return value
> Root site Url of the given site Url

#### NormalizedTimerJobName(System.String)
Normalizes the timer job name
> ##### Parameters
> **timerJobName:** Timer job name

> ##### Return value
> Normalized timer job name

#### IsInternalServerErrorException(System.Exception)
Returns true if the exception was a "The remote server returned an error: (500) Internal Server Error"
> ##### Parameters
> **ex:** Exception to examine

> ##### Return value
> True if "The remote server returned an error: (500) Internal Server Error" exception, false otherwise

#### IsNotFoundException(System.Exception)
Returns true if the exception was a "The remote server returned an error: (404) Not Found"
> ##### Parameters
> **ex:** Exception to examine

> ##### Return value
> True if "The remote server returned an error: (404) Not Found" exception, false otherwise

#### Constructor
Constructor used when state is being managed by the timer job framework
> ##### Parameters
> **url:** Url of the site the timer job is operating against

> **siteClientContext:** ClientContext object for the root site of the site collection

> **webClientContext:** ClientContext object for passed site Url

> **tenantClientContext:** ClientContext object to work with the Tenant API

> **previousRun:** Datetime of the last run

> **previousRunSuccessful:** Bool showing if the previous run was successful

> **previousRunVersion:** Version of the timer job that was used for the previous run

> **properties:** Custom keyword value collection that can be used to persist custom properties

> **configurationData:** Optional timerjob configuration data


#### Constructor
Constructor used when state is not managed
> ##### Parameters
> **url:** Url of the site the timer job is operating against

> **ccSite:** ClientContext object for the root site of the site collection

> **ccWeb:** ClientContext object for passed site Url

> **ccTenant:** Tenant ClientContext

> **configurationData:** Configuration data


#### 
Gets a property from the custom properties list
> ##### Parameters
> **propertyKey:** Key of the property to retrieve

> ##### Return value
> Value of the requested property

#### 
Adds or updates a property in the custom properties list
> ##### Parameters
> **propertyKey:** Key of the property to add or update

> **propertyValue:** Value of the property to add or update


#### 
Deletes a property from the custom property list
> ##### Parameters
> **propertyKey:** Name of the property to delete


## Core.Framework.TimerJobs.TimerJobRun
            
Class that holds the state information that's being stored in the web property bag of web that's being "processed"
        
### Fields

#### PreviousRun
DateTime of the previous run attempt
#### PreviousRunSuccessful
Bool indicating if the previous run was successful
#### PreviousRunVersion
Timer job version used during the previous run
#### Properties
Property value collection used to store timer job custom properties
### Methods


#### Constructor
Constructor used when state is being managed by the timer job framework
> ##### Parameters
> **url:** Url of the site the timer job is operating against

> **siteClientContext:** ClientContext object for the root site of the site collection

> **webClientContext:** ClientContext object for passed site Url

> **tenantClientContext:** ClientContext object to work with the Tenant API

> **previousRun:** Datetime of the last run

> **previousRunSuccessful:** Bool showing if the previous run was successful

> **previousRunVersion:** Version of the timer job that was used for the previous run

> **properties:** Custom keyword value collection that can be used to persist custom properties

> **configurationData:** Optional timerjob configuration data


#### Constructor
Constructor used when state is not managed
> ##### Parameters
> **url:** Url of the site the timer job is operating against

> **ccSite:** ClientContext object for the root site of the site collection

> **ccWeb:** ClientContext object for passed site Url

> **ccTenant:** Tenant ClientContext

> **configurationData:** Configuration data


#### 
Gets a property from the custom properties list
> ##### Parameters
> **propertyKey:** Key of the property to retrieve

> ##### Return value
> Value of the requested property

#### 
Adds or updates a property in the custom properties list
> ##### Parameters
> **propertyKey:** Key of the property to add or update

> **propertyValue:** Value of the property to add or update


#### 
Deletes a property from the custom property list
> ##### Parameters
> **propertyKey:** Name of the property to delete


## Core.Framework.TimerJobs.TimerJobRunEventArgs
            
Event arguments for the TimerJobRun event
        
### Methods


#### Constructor
Constructor used when state is being managed by the timer job framework
> ##### Parameters
> **url:** Url of the site the timer job is operating against

> **siteClientContext:** ClientContext object for the root site of the site collection

> **webClientContext:** ClientContext object for passed site Url

> **tenantClientContext:** ClientContext object to work with the Tenant API

> **previousRun:** Datetime of the last run

> **previousRunSuccessful:** Bool showing if the previous run was successful

> **previousRunVersion:** Version of the timer job that was used for the previous run

> **properties:** Custom keyword value collection that can be used to persist custom properties

> **configurationData:** Optional timerjob configuration data


#### Constructor
Constructor used when state is not managed
> ##### Parameters
> **url:** Url of the site the timer job is operating against

> **ccSite:** ClientContext object for the root site of the site collection

> **ccWeb:** ClientContext object for passed site Url

> **ccTenant:** Tenant ClientContext

> **configurationData:** Configuration data


#### GetProperty(System.String)
Gets a property from the custom properties list
> ##### Parameters
> **propertyKey:** Key of the property to retrieve

> ##### Return value
> Value of the requested property

#### SetProperty(System.String,System.String)
Adds or updates a property in the custom properties list
> ##### Parameters
> **propertyKey:** Key of the property to add or update

> **propertyValue:** Value of the property to add or update


#### DeleteProperty(System.String)
Deletes a property from the custom property list
> ##### Parameters
> **propertyKey:** Name of the property to delete


## Core.Framework.TimerJobs.Utilities.SiteEnumeration
            
Singleton class that's responsible for resolving wildcard site Url's into a list af site Url's
        
### Properties

#### Instance
Singleton instance to access this class
### Methods


#### ResolveSite(Microsoft.Online.SharePoint.TenantAdministration.Tenant,System.String,System.Collections.Generic.List{System.String})
Builds up a list of site collections that match the passed site wildcard. This method can be used against Office 365
> ##### Parameters
> **tenant:** Tenant object to use for resolving the regular sites

> **siteWildCard:** The widcard site Url (e.g. https://tenant.sharepoint.com/sites/*)

> **resolvedSites:** List of site collections matching the passed wildcard site Url


#### ResolveSite(Microsoft.SharePoint.Client.ClientContext,System.String,System.Collections.Generic.List{System.String})
Builds up a list of site collections that match the passed site wildcard. This method can be used against on-premises
> ##### Parameters
> **context:** ClientContext object of an arbitrary site collection accessible by the defined enumeration username and password

> **site:** The widcard site Url (e.g. https://tenant.sharepoint.com/sites/*)

> **resolvedSites:** List of site collections matching the passed wildcard site Url


#### FillSitesViaTenantAPIAndSearch(Microsoft.Online.SharePoint.TenantAdministration.Tenant)
Fill site list via tenant API for "regular" site collections. Search API is used for OneDrive for Business site collections
> ##### Parameters
> **tenant:** Tenant object to operate against


#### FillSitesViaSearch(Microsoft.SharePoint.Client.ClientContext)
Fill site list via the Search API. Applies to all type of sites. Typically used in on-premises environments
> ##### Parameters
> **context:** ClientContext object of an arbitrary site collection accessible by the defined enumeration username and password


#### SiteSearch(Microsoft.SharePoint.Client.ClientRuntimeContext,System.String)
Get all sites that match the passed query. Batching is done in batches of 500 as this is compliant for both Office 365 as SharePoint on-premises
> ##### Parameters
> **cc:** ClientContext object of an arbitrary site collection accessible by the defined enumeration username and password

> **keywordQueryValue:** Query string

> ##### Return value
> List of found site collections

#### ProcessQuery(Microsoft.SharePoint.Client.ClientRuntimeContext,System.String,System.Collections.Generic.List{System.String},Microsoft.SharePoint.Client.Search.Query.KeywordQuery,System.Int32)

> ##### Parameters
> **cc:** ClientContext object of an arbitrary site collection accessible by the defined enumeration username and password

> **keywordQueryValue:** Query to execute

> **sites:** List of found site collections

> **keywordQuery:** KeywordQuery instance that will perform the actual queries

> **startRow:** Row as of which we want to see the results

> ##### Return value
> Total result rows of the query

## Core.Pages.DefaultClientSideWebParts
            
List of possible OOB web parts
        

## Core.Pages.ClientSidePage
            
Represents a modern client side page with all it's contents
        

## Core.Pages.CanvasZoneTemplate
            
The type of canvas being used
        

## Core.Pages.CanvasZone
            
Represents a zone on the canvas
        

## Core.Pages.CanvasSection
            
Represents a section in a canvas zone
        

## Core.Pages.AvailableClientSideComponents
            
Class holding a collection of client side webparts (retrieved via the _api/web/GetClientSideWebParts REST call)
        

## Core.Pages.ClientSideComponent
            
Client side webpart object (retrieved via the _api/web/GetClientSideWebParts REST call)
        

## Core.Pages.CanvasControl
            
Base control
        
### Methods


#### ToHtml
Converts a control object to it's html representation
> ##### Return value
> Html representation of a control

#### Delete
Removes the control from the page

#### GetType(System.String)
Receives data-sp-controldata content and detects the type of the control
> ##### Parameters
> **controlDataJson:** data-sp-controldata json string

> ##### Return value
> Type of the control represented by the json string

## Core.Pages.ClientSideText
            
Controls of type 4 ( = text control)
        
### Methods


#### ToHtml
Converts this control to it's html representation
> ##### Return value
> Html representation of this control

## Core.Pages.ClientSideWebPart
            
This class is used to instantiate controls of type 3 (= client side web parts). Using this class you can instantiate a control and add it on a .
        
### Properties

#### JsonWebPartData
Json serialized web part properties
#### HtmlPropertiesData
Html properties data
#### HtmlProperties
Value of the data-sp-htmlproperties attribute
#### WebPartId
ID of the client side web part
#### WebPartData
Value of the data-sp-webpart attribute
#### Title
Title of the web part
#### Description
Description of the web part
#### PropertiesJson
Json serialized web part properties
#### Properties
Web properties as configurable
#### Type
Return of the client side web part
### Methods


#### Constructor
Instantiates client side web part from scratch.

#### Constructor
Instantiates a client side web part based on the information that was obtain from calling the AvailableClientSideComponents methods on the object.
> ##### Parameters
> **component:** Component to create a ClientSideWebPart instance for


#### Import(OfficeDevPnP.Core.Pages.ClientSideComponent,System.Func{System.String,System.String})
Imports a to use it as base for configuring the client side web part instance
> ##### Parameters
> **component:** to import

> **clientSideWebPartPropertiesUpdater:** Function callback that allows you to manipulate the client side web part properties after import


#### ToHtml
Returns a HTML representation of the client side web part
> ##### Return value
> HTML representation of the client side web part

## Core.Pages.ClientSideCanvasControlData
            
Abstract base class representing the json control data that will be included in each client side control (de-)serialization (data-sp-controldata attribute)
        

## Core.Pages.ClientSideTextControlData
            
Control data for controls of type 4 (= text control)
        

## Core.Pages.ClientSideWebPartControlData
            
Control data for controls of type 3 (= client side web parts)
        

## Core.Pages.ClientSideWebPartData
            
Json web part data that will be included in each client side web part (de-)serialization (data-sp-webpartdata)
        

## Core.Entities.UnifiedGroupEntity
            
Defines a Unified Group
        

## Core.Entities.AreaNavigationEntity
            
Entity description navigation
        
### Properties

#### GlobalNavigation
Specifies the Global Navigation (top bar navigation)
#### CurrentNavigation
Specifies the Current Navigation (quick launch navigation)
#### Sorting
Defines the sorting
#### SortAscending
Defines if sorted ascending
#### SortBy
Defines sorted by value
### Methods


#### Constructor
ctor

## Core.Entities.CustomActionEntity
            
CustomActionEntity class describes the information for a Custom Action
        
### Properties

#### CommandUIExtension
Gets or sets a value that specifies an implementation specific XML fragment that determines user interface properties of the custom action
#### RegistrationId
Gets or sets the value that specifies the identifier of the object associated with the custom action.
#### RegistrationType
Specifies the type of object associated with the custom action. A Nullable Type
#### Name
Gets or sets the name of the custom action.
#### Description
Description of the custom action
#### Title
Custom action title
#### Location
Custom action location (A string that contains the location; for example, Microsoft.SharePoint.SiteSettings)
#### ScriptBlock
The JavaScript to be executed by this custom action
#### Sequence
The sequence number of this custom action
#### ImageUrl
The URL to the image used for this custom action
#### Group
The group of this custom action
#### Url
The URL this custom action should navigate the user to
#### Rights
Gets or sets the value that specifies the permissions needed for the custom action.
#### Remove
Indicates if the custom action will need to be removed
#### ScriptSrc
Gets or sets a value that specifies the URI of a file which contains the ECMAScript to execute on the page

## Core.Entities.DefaultColumnTermPathValue
            
Specifies a default column value for a document library
        
### Properties

#### FolderRelativePath
The Path of the folder, Rootfolder of the document library is "/"
#### FieldInternalName
The internal name of the field
#### TermPaths
Taxonomy paths in the shape of "TermGroup|TermSet|Term"
### Methods


#### Constructor
ctor

## Core.Entities.IDefaultColumnValue
            
IDefaultColumnValue
        
### Properties

#### FolderRelativePath
Folder relative path
#### FieldInternalName
Field internal name

## Core.Entities.DefaultColumnTermValue
            
Specifies a default column value for a document library
        
### Properties

#### Terms
Taxonomy paths in the shape of "TermGroup|TermSet|Term"

## Core.Entities.DefaultColumnTextValue
            
DefaultColumnTextValue
        

## Core.Entities.DefaultColumnValue
            
DefaultColumnValue
        
### Properties

#### FolderRelativePath
The Path of the folder, Rootfolder of the document library is "/"
#### FieldInternalName
The internal name of the field

## Core.Entities.ExternalUserEntity
            
External user entity
        

## Core.Entities.SiteEntity
            
SiteEntity class describes the information for a SharePoint site (collection)
        
### Properties

#### Url
The SPO url
#### Title
The site title
#### Description
The site description
#### SiteOwnerLogin
The site owner
#### CurrentResourceUsage
The current resource usage points
#### IndexDocId
IndexDocId for Search Paging
#### Lcid
The site locale. See http://technet.microsoft.com/en-us/library/ff463597.aspx for a complete list of Lcid's
#### StorageMaximumLevel
Site quota in MB
#### StorageUsage
The storage quota usage in MB
#### StorageWarningLevel
Site quota warning level in MB
#### LastContentModifiedDate
The last modified date/time of the site collection's content
#### Template
Site template being used
#### TimeZoneId
TimeZoneID for the site. "(UTC+01:00) Brussels, Copenhagen, Madrid, Paris" = 3 See http://blog.jussipalo.com/2013/10/list-of-sharepoint-timezoneid-values.html for a complete list
#### UserCodeMaximumLevel
The user code quota in points
#### UserCodeWarningLevel
The user code quota warning level in points
#### WebsCount
The count of the SPWeb objects in the site collection

## Core.Entities.SitePolicyEntity
            
Properties of a site policy object
        
### Properties

#### Description
The description of the policy
#### EmailBody
The body of the notification email if there is no site mailbox associated with the site.
#### EmailBodyWithTeamMailbox
The body of the notification email if there is a site mailbox associated with the site.
#### EmailSubject
The subject of the notification email.
#### Name
The name of the policy

## Core.Entities.VariationInformation
            
Class containing variation configuration information
        
### Properties

#### AutomaticCreation
Automatic creation Mapped to property "EnableAutoSpawnPropertyName"
#### RecreateDeletedTargetPage
Recreate Deleted Target Page; set to false to enable recreation Mapped to property "AutoSpawnStopAfterDeletePropertyName"
#### UpdateTargetPageWebParts
Update Target Page Web Parts Mapped to property "UpdateWebPartsPropertyName"
#### CopyResources
Copy resources Mapped to property "CopyResourcesPropertyName"
#### SendNotificationEmail
Send email notification Mapped to property "SendNotificationEmailPropertyName"
#### RootWebTemplate
Configuration setting site template to be used for the top sites of each label Mapped to property "SourceVarRootWebTemplatePropertyName"

## Core.Entities.VariationLabelEntity
            
Class represents variation label
        
### Properties

#### Title
The variation label title
#### Description
The variation label description
#### FlagControlDisplayName
The flag to control display name
#### Language
The variation label language
#### Locale
The variation label locale
#### HierarchyCreationMode
The hierarchy creation mode
#### IsSource
Set as source variation
#### IsCreated
Gets a value indicating whether the variation label has been created

## Core.Entities.WebPartEntity
            
Class that describes information about a web part
        
### Properties

#### WebPartXml
XML representation of the web part
#### WebPartZone
Zone that will need to hold the web part
#### WebPartIndex
Index (order) of the web part in it's zone
#### WebPartTitle
Title of the web part

## Core.Entities.YammerGroup
            
Represents Yammer Group information Generated based on Yammer response on 30th of June 2014 and using http://json2csharp.com/ service
        

## Core.Entities.YammerUser
            
Represents YammerUser Generated based on Yammer response on 30th of June 2014 and using http://json2csharp.com/ service
        

## Core.AzureEnvironment
            
Enum to identify the supported Office 365 hosting environments
        

## Core.AuthenticationManager
            
This manager class can be used to obtain a SharePointContext object
            
        
### Methods


#### GetSharePointOnlineAuthenticatedContextTenant(System.String,System.String,System.String)
Returns a SharePointOnline ClientContext object
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **tenantUser:** User to be used to instantiate the ClientContext object

> **tenantUserPassword:** Password of the user used to instantiate the ClientContext object

> ##### Return value
> ClientContext to be used by CSOM code

#### GetSharePointOnlineAuthenticatedContextTenant(System.String,System.String,System.Security.SecureString)
Returns a SharePointOnline ClientContext object
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **tenantUser:** User to be used to instantiate the ClientContext object

> **tenantUserPassword:** Password (SecureString) of the user used to instantiate the ClientContext object

> ##### Return value
> ClientContext to be used by CSOM code

#### GetAppOnlyAuthenticatedContext(System.String,System.String,System.String)
Returns an app only ClientContext object
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **appId:** Application ID which is requesting the ClientContext object

> **appSecret:** Application secret of the Application which is requesting the ClientContext object

> ##### Return value
> ClientContext to be used by CSOM code

#### GetAppOnlyAuthenticatedContext(System.String,System.String,System.String,OfficeDevPnP.Core.AzureEnvironment)
Returns an app only ClientContext object
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **appId:** Application ID which is requesting the ClientContext object

> **appSecret:** Application secret of the Application which is requesting the ClientContext object

> **environment:** SharePoint environment being used

> ##### Return value
> ClientContext to be used by CSOM code

#### GetAppOnlyAuthenticatedContext(System.String,System.String,System.String,System.String,System.String,System.String)
Returns an app only ClientContext object
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **realm:** Realm of the environment (tenant) that requests the ClientContext object

> **appId:** Application ID which is requesting the ClientContext object

> **appSecret:** Application secret of the Application which is requesting the ClientContext object

> **acsHostUrl:** Azure ACS host, defaults to accesscontrol.windows.net but internal pre-production environments use other hosts

> **globalEndPointPrefix:** Azure ACS endpoint prefix, defaults to accounts but internal pre-production environments use other prefixes

> ##### Return value
> ClientContext to be used by CSOM code

#### GetAzureADACSEndPoint(OfficeDevPnP.Core.AzureEnvironment)
Get's the Azure ASC login end point for the given environment
> ##### Parameters
> **environment:** Environment to get the login information for

> ##### Return value
> Azure ASC login endpoint

#### GetAzureADACSEndPointPrefix(OfficeDevPnP.Core.AzureEnvironment)
Get's the Azure ACS login end point prefix for the given environment
> ##### Parameters
> **environment:** Environment to get the login information for

> ##### Return value
> Azure ACS login endpoint prefix

#### EnsureToken(System.String,System.String,System.String,System.String,System.String,System.String)
Ensure that AppAccessToken is filled with a valid string representation of the OAuth AccessToken. This method will launch handle with token cleanup after the token expires
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **realm:** Realm of the environment (tenant) that requests the ClientContext object

> **appId:** Application ID which is requesting the ClientContext object

> **appSecret:** Application secret of the Application which is requesting the ClientContext object

> **acsHostUrl:** Azure ACS host, defaults to accesscontrol.windows.net but internal pre-production environments use other hosts

> **globalEndPointPrefix:** Azure ACS endpoint prefix, defaults to accounts but internal pre-production environments use other prefixes


#### GetAccessTokenLease(System.DateTime)
Get the access token lease time span.
> ##### Parameters
> **expiresOn:** The ExpiresOn time of the current access token

> ##### Return value
> Returns a TimeSpan represents the time interval within which the current access token is valid thru.

#### GetWebLoginClientContext(System.String,System.Drawing.Icon)
Returns a SharePoint on-premises / SharePoint Online ClientContext object. Requires claims based authentication with FedAuth cookie.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **icon:** Optional icon to use for the popup form

> ##### Return value
> ClientContext to be used by CSOM code

#### GetNetworkCredentialAuthenticatedContext(System.String,System.String,System.String,System.String)
Returns a SharePoint on-premises / SharePoint Online Dedicated ClientContext object
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **user:** User to be used to instantiate the ClientContext object

> **password:** Password of the user used to instantiate the ClientContext object

> **domain:** Domain of the user used to instantiate the ClientContext object

> ##### Return value
> ClientContext to be used by CSOM code

#### GetNetworkCredentialAuthenticatedContext(System.String,System.String,System.Security.SecureString,System.String)
Returns a SharePoint on-premises / SharePoint Online Dedicated ClientContext object
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **user:** User to be used to instantiate the ClientContext object

> **password:** Password (SecureString) of the user used to instantiate the ClientContext object

> **domain:** Domain of the user used to instantiate the ClientContext object

> ##### Return value
> ClientContext to be used by CSOM code

#### GetHighTrustCertificateAppOnlyAuthenticatedContext(System.String,System.String,System.String,System.String,System.String)
Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The SharePoint Client ID

> **certificatePath:** Full path to the private key certificate (.pfx) used to authenticate

> **certificatePassword:** Password used for the private key certificate (.pfx)

> **certificateIssuerId:** The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer

> ##### Return value
> Authenticated SharePoint ClientContext

#### GetHighTrustCertificateAppOnlyAuthenticatedContext(System.String,System.String,System.String,System.Security.SecureString,System.String)
Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The SharePoint Client ID

> **certificatePath:** Full path to the private key certificate (.pfx) used to authenticate

> **certificatePassword:** Password used for the private key certificate (.pfx)

> **certificateIssuerId:** The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer

> ##### Return value
> Authenticated SharePoint ClientContext

#### GetHighTrustCertificateAppOnlyAuthenticatedContext(System.String,System.String,System.Security.Cryptography.X509Certificates.StoreName,System.Security.Cryptography.X509Certificates.StoreLocation,System.String,System.String)
Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The SharePoint Client ID

> **storeName:** The name of the store for the certificate

> **storeLocation:** The location of the store for the certificate

> **thumbPrint:** The thumbprint of the certificate to locate in the store

> **certificateIssuerId:** The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer

> ##### Return value
> Authenticated SharePoint ClientContext

#### GetHighTrustCertificateAppOnlyAuthenticatedContext(System.String,System.String,System.Security.Cryptography.X509Certificates.X509Certificate2,System.String)
Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The SharePoint Client ID

> **certificate:** Private key certificate (.pfx) used to authenticate

> **certificateIssuerId:** The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer

> ##### Return value
> Authenticated SharePoint ClientContext

#### GetAzureADNativeApplicationAuthenticatedContext(System.String,System.String,System.String,Microsoft.IdentityModel.Clients.ActiveDirectory.TokenCache,OfficeDevPnP.Core.AzureEnvironment)
Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The Azure AD Native Application Client ID

> **redirectUrl:** The Azure AD Native Application Redirect Uri as a string

> **tokenCache:** Optional token cache. If not specified an in-memory token cache will be used

> **environment:** SharePoint environment being used

> ##### Return value
> Client context object

#### GetAzureADNativeApplicationAuthenticatedContext(System.String,System.String,System.Uri,Microsoft.IdentityModel.Clients.ActiveDirectory.TokenCache,OfficeDevPnP.Core.AzureEnvironment)
Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The Azure AD Native Application Client ID

> **redirectUri:** The Azure AD Native Application Redirect Uri

> **tokenCache:** Optional token cache. If not specified an in-memory token cache will be used

> **environment:** SharePoint environment being used

> ##### Return value
> Client context object

#### GetAzureADWebApplicationAuthenticatedContext(System.String,System.Func{System.String,System.String})
Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Web Application registered. The user will not be prompted for authentication, the current user's authentication context will be used by leveraging ADAL.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **accessTokenGetter:** The AccessToken getter method to use

> ##### Return value
> Client context object

#### GetAzureADAccessTokenAuthenticatedContext(System.String,System.String)
Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Web Application registered. The user will not be prompted for authentication, the current user's authentication context will be used by leveraging an explicit OAuth 2.0 Access Token value.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **accessToken:** An explicit value for the AccessToken

> ##### Return value
> Client context object

#### GetAzureADAppOnlyAuthenticatedContext(System.String,System.String,System.String,System.Security.Cryptography.X509Certificates.StoreName,System.Security.Cryptography.X509Certificates.StoreLocation,System.String,OfficeDevPnP.Core.AzureEnvironment)
Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The Azure AD Application Client ID

> **tenant:** The Azure AD Tenant, e.g. mycompany.onmicrosoft.com

> **storeName:** The name of the store for the certificate

> **storeLocation:** The location of the store for the certificate

> **thumbPrint:** The thumbprint of the certificate to locate in the store

> **environment:** SharePoint environment being used

> ##### Return value
> ClientContext being used

#### GetAzureADAppOnlyAuthenticatedContext(System.String,System.String,System.String,System.String,System.String,OfficeDevPnP.Core.AzureEnvironment)
Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The Azure AD Application Client ID

> **tenant:** The Azure AD Tenant, e.g. mycompany.onmicrosoft.com

> **certificatePath:** The path to the certificate (*.pfx) file on the file system

> **certificatePassword:** Password to the certificate

> **environment:** SharePoint environment being used

> ##### Return value
> Client context object

#### GetAzureADAppOnlyAuthenticatedContext(System.String,System.String,System.String,System.String,System.Security.SecureString,OfficeDevPnP.Core.AzureEnvironment)
Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The Azure AD Application Client ID

> **tenant:** The Azure AD Tenant, e.g. mycompany.onmicrosoft.com

> **certificatePath:** The path to the certificate (*.pfx) file on the file system

> **certificatePassword:** Password to the certificate

> **environment:** SharePoint environment being used

> ##### Return value
> Client context object

#### GetAzureADAppOnlyAuthenticatedContext(System.String,System.String,System.String,System.Security.Cryptography.X509Certificates.X509Certificate2,OfficeDevPnP.Core.AzureEnvironment)
Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
> ##### Parameters
> **siteUrl:** Site for which the ClientContext object will be instantiated

> **clientId:** The Azure AD Application Client ID

> **tenant:** The Azure AD Tenant, e.g. mycompany.onmicrosoft.com

> **certificate:** Certificate used to authenticate

> **environment:** SharePoint environment being used

> ##### Return value
> 

#### GetAzureADLoginEndPoint(OfficeDevPnP.Core.AzureEnvironment)
Get's the Azure AD login end point for the given environment
> ##### Parameters
> **environment:** Environment to get the login information for

> ##### Return value
> Azure AD login endpoint

#### GetADFSUserNameMixedAuthenticatedContext(System.String,System.String,System.String,System.String,System.String,System.String,System.Int32)
Returns a SharePoint on-premises ClientContext for sites secured via ADFS
> ##### Parameters
> **siteUrl:** Url of the SharePoint site that's secured via ADFS

> **user:** Name of the user (e.g. administrator)

> **password:** Password of the user

> **domain:** Windows domain of the user

> **sts:** Hostname of the ADFS server (e.g. sts.company.com)

> **idpId:** Identifier of the ADFS relying party that we're hitting

> **logonTokenCacheExpirationWindow:** Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.

> ##### Return value
> ClientContext to be used by CSOM code

#### RefreshADFSUserNameMixedAuthenticatedContext(System.String,System.String,System.String,System.String,System.String,System.String,System.Int32)
Refreshes the SharePoint FedAuth cookie
> ##### Parameters
> **siteUrl:** Url of the SharePoint site that's secured via ADFS

> **user:** Name of the user (e.g. administrator)

> **password:** Password of the user

> **domain:** Windows domain of the user

> **sts:** Hostname of the ADFS server (e.g. sts.company.com)

> **idpId:** Identifier of the ADFS relying party that we're hitting

> **logonTokenCacheExpirationWindow:** Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.


#### GetADFSCertificateMixedAuthenticationContext(System.String,System.String,System.String,System.String,System.Int32)
Returns a SharePoint on-premises ClientContext for sites secured via ADFS
> ##### Parameters
> **siteUrl:** Url of the SharePoint site that's secured via ADFS

> **serialNumber:** Represents the serial number of the certificate as displayed by the certificate dialog box, but without the spaces, or as returned by the System.Security.Cryptography.X509Certificates.X509Certificate.GetSerialNumberString method

> **sts:** Hostname of the ADFS server (e.g. sts.company.com)

> **idpId:** Identifier of the ADFS relying party that we're hitting

> **logonTokenCacheExpirationWindow:** Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.

> ##### Return value
> ClientContext to be used by CSOM code

#### RefreshADFSCertificateMixedAuthenticationContext(System.String,System.String,System.String,System.String,System.Int32)
Refreshes the SharePoint FedAuth cookie
> ##### Parameters
> **siteUrl:** Url of the SharePoint site that's secured via ADFS

> **serialNumber:** Represents the serial number of the certificate as displayed by the certificate dialog box, but without the spaces, or as returned by the System.Security.Cryptography.X509Certificates.X509Certificate.GetSerialNumberString method

> **sts:** Hostname of the ADFS server (e.g. sts.company.com)

> **idpId:** Identifier of the ADFS relying party that we're hitting

> **logonTokenCacheExpirationWindow:** Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.

> ##### Return value
> ClientContext to be used by CSOM code

## Core.BuiltInContentTypeId
            
A class that returns strings that represent identifiers (IDs) for built-in content types.
        
### Fields

#### DocumentSet
Contains the content identifier (ID) for the DocumentSet content type. To get content type from a list, use BestMatchContentTypeId().
#### Item
Contains the content identifier (ID) for the Item content type.
#### ReusableHTML
Contains the content identifier (ID) for content types used in the publishing infrastructure.
#### ModernArticlePage
Contains the content identifier (ID) for content types used in the modern page infrastructure

## Core.Constants
            
Constants.
            Recommendation: Constants should follow C# style guidelines and be Pascal Case
        

## Core.Diagnostics.Log
            
Logging class
        

## Core.Enums.NavigationType
            
Enum that defines the navigation types
        

## Core.Enums.TimeZone
            
Timezones to use when creating sitecollections Format UTC[PLUS|MINUS][HH:MM]_[DESCRIPTION]
        

## Core.EcmListManualRecordDeclaration
            
Specifies whether this list should allow the manual declaration of records. When manual record declaration is unavailable, records can only be declared through a policy or workflow.
        
### Fields

#### UseSiteCollectionDefaults
Use the site collection defaults
#### AlwaysAllowManualDeclaration
Always allow to manual declare records in this list
#### NeverAllowManualDeclaration
Never allow to manual declare records in this list

## Core.EcmRecordDeclarationBy
            
Specify which user roles can declare and undeclare record status manually
        
### Fields

#### AllListContributors
All list contributors and administrators
#### OnlyAdmins
Only list administrators
#### OnlyPolicy
Only policy actions

## Core.EcmSiteRecordRestrictions
            
Specify restrictions to place on a document or item once it has been declared as a record. Changing this setting will not affect items which have already been declared records.
        
### Fields

#### None
Records are no more restricted than non-records
#### BlockDelete
Records can be edited but not deleted
#### BlockEdit
Records cannot be edited or deleted. Any change will require the record declaration to be revoked

## Core.VotingExperience
            
Voting Experience in List
        

## Core.WikiPageLayout
            
Out of the box wiki page layouts enumeration
        

## Core.Extensions.DictionaryExtensions
            
Extension type for Dictionaries
        

## Core.Extensions.EnumerableExtensions
            
Extension methods to make working with IEnumerable<T> values easier.
        
### Methods


#### DeepEquals``1(System.Collections.Generic.IEnumerable{``0},System.Collections.Generic.IEnumerable{``0})
Compares to instances of IEnumerable<T>
> ##### Parameters
> **source:** Source enumeration

> **target:** Target enumeration

> ##### Return value
> Wether the two enumerations are deep equal

## Core.IdentityModel.TokenProviders.ADFS.CertificateMixed
            
ADFS Active authentication based on username + password. Uses the trust/13/usernamemixed ADFS endpoint.
        
### Methods


#### GetFedAuthCookie(System.String,System.String,System.Uri,System.String,System.Int32)
Performs active authentication against ADFS using the trust/13/usernamemixed ADFS endpoint.
> ##### Parameters
> **siteUrl:** Url of the SharePoint site that's secured via ADFS

> **serialNumber:** Serial Number of the Current User > My Certificate to use to authenticate

> **certificateMixed:** Uri to the ADFS certificatemixed endpoint

> **relyingPartyIdentifier:** Identifier of the ADFS relying party that we're hitting

> **logonTokenCacheExpirationWindow:** 

> ##### Return value
> A cookiecontainer holding the FedAuth cookie

#### RequestToken(System.String,System.Uri,System.String)
Returns Generic XML Security Token from ADFS to generated FedAuth
> ##### Parameters
> **serialNumber:** Serial Number of Certificate from CurrentUSer > My Certificate

> **certificateMixed:** ADFS Endpoint for Certificate Mixed Authentication

> **relyingPartyIdentifier:** Identifier of the ADFS relying party that we're hitting

> ##### Return value
> 

## Core.IdentityModel.TokenProviders.ADFS.BaseProvider
            
Base class for active SAML based authentication
        
### Methods


#### TransformSamlTokenToFedAuth(System.String,System.String,System.String)
Transforms the retrieved SAML token into a FedAuth cookie value by calling into the SharePoint STS
> ##### Parameters
> **samlToken:** SAML token obtained via active authentication to ADFS

> **samlSite:** Url of the SAML secured SharePoint site

> **relyingPartyIdentifier:** Identifier of the ADFS relying party that we're hitting

> ##### Return value
> The FedAuth cookie value

#### WrapInSoapMessage(System.String,System.String)
Wrap SAML token in RequestSecurityTokenResponse soap message
> ##### Parameters
> **stsResponse:** SAML token obtained via active authentication to ADFS

> **relyingPartyIdentifier:** Identifier of the ADFS relying party that we're hitting

> ##### Return value
> RequestSecurityTokenResponse soap message

#### SamlTokenExpiresOn(System.String)
Returns the DateTime when then received saml token will expire
> ##### Parameters
> **stsResponse:** saml token

> ##### Return value
> DateTime holding the expiration date. Defaults to DateTime.MinValue if there's no valid datetime in the saml token

#### SamlTokenlifeTime(System.String)
Returns the SAML token life time
> ##### Parameters
> **stsResponse:** saml token

> ##### Return value
> TimeSpan holding the token lifetime. Defaults to TimeSpan.Zero is case of problems

## Core.IdentityModel.TokenProviders.ADFS.UsernameMixed
            
ADFS Active authentication based on username + password. Uses the trust/13/usernamemixed ADFS endpoint.
        
### Methods


#### GetFedAuthCookie(System.String,System.String,System.String,System.Uri,System.String,System.Int32)
Performs active authentication against ADFS using the trust/13/usernamemixed ADFS endpoint.
> ##### Parameters
> **siteUrl:** Url of the SharePoint site that's secured via ADFS

> **userName:** Name of the user (e.g. domain\administrator)

> **password:** Password of th user

> **userNameMixed:** Uri to the ADFS usernamemixed endpoint

> **relyingPartyIdentifier:** Identifier of the ADFS relying party that we're hitting

> **logonTokenCacheExpirationWindow:** 

> ##### Return value
> A cookiecontainer holding the FedAuth cookie

## Core.Utilities.CookieReader
            
WinInet.dll wrapper
        
### Fields

#### INTERNET_COOKIE_HTTPONLY
Enables the retrieval of cookies that are marked as "HTTPOnly". Do not use this flag if you expose a scriptable interface, because this has security implications. It is imperative that you use this flag only if you can guarantee that you will never expose the cookie to third-party code by way of an extensibility mechanism you provide. Version: Requires Internet Explorer 8.0 or later.
### Methods


#### GetCookie(System.String)
Returns cookie contents as a string
> ##### Parameters
> **url:** 

> ##### Return value
> 

## Core.Utilities.WebParts.Processors.PassThroughProcessor
            
Default processor when others are not resolved
        

## Core.Utilities.WebParts.Processors.XsltWebPartPostProcessor
            
Updates view for XsltListViewWebPart using schema definition provided Instead of using default view for XsltListViewWebPart, it tries to resolve view from schema and updates hidden view created by XsltListViewWebPart
        

## Core.Utilities.WebParts.Schema.WebParts
            

        
### Properties

#### WebPart


## Core.Utilities.WebParts.Schema.WebPart
            

        
### Properties

#### 

#### MetaData

#### Data

#### 

#### 

#### 

#### 

#### 

#### 


## Core.Utilities.WebParts.Schema.WebPartMetaData
            

        
### Properties

#### Type

#### ImportErrorMessage

#### 

#### 


## Core.Utilities.WebParts.Schema.WebPartMetaDataType
            

        
### Properties

#### Name

#### Src


## Core.Utilities.WebParts.Schema.WebPartData
            

        
### Properties

#### Properties

#### GenericWebPartProperties


## Core.Utilities.WebParts.Schema.PropertyContainerType
            

        
### Properties

#### Property

#### PropertySpecified
Gets a value indicating whether the Property collection is empty.
#### Ipersonalizable

#### IpersonalizableSpecified
Gets a value indicating whether the Ipersonalizable collection is empty.
#### 

#### 
Gets a value indicating whether the Property collection is empty.
### Methods


#### Constructor
Initializes a new instance of the class.

#### Constructor
Initializes a new instance of the class.

## Core.Utilities.WebParts.Schema.PropertyType
            

        
### Properties

#### Value
Gets or sets the text value.
#### Name

#### Type

#### Null

#### NullSpecified
Gets or sets a value indicating whether the Null property is specified.

## Core.Utilities.WebParts.Schema.PropertyContainerTypeIpersonalizable
            

        
### Properties

#### Property

#### PropertySpecified
Gets a value indicating whether the Property collection is empty.
### Methods


#### Constructor
Initializes a new instance of the class.

## Core.Utilities.WebParts.WebPartPostProcessorFactory
            
Creates by parsing web part schema xml
        

## Core.Utilities.CAML
            
Use this class to build your CAML xml and avoid XML issues.
            
            CAML.ViewQuery(
                CAML.Where(
                    CAML.And(
                        CAML.Eq(CAML.FieldValue("Project", "Integer", "{0}")),
                        CAML.Geq(CAML.FieldValue("StartDate","DateTime", CAML.Today()))
                    )
                ),
                CAML.OrderBy(
                    new OrderByField("StartDate", false),
                    new OrderByField("Title")
                ),
                rowLimit: 5
            );
            
        
### Methods


#### Today(System.Nullable{System.Int32})
Creates the <Today /> node.
> ##### Parameters
> **offset:** Time offset from today (+5 days or -5 days, for example).

> ##### Return value
> 

#### ViewQuery(System.String,System.String,System.Int32)
Root <View> and <Query> nodes.
> ##### Parameters
> **whereClause:** <Where> node.

> **orderByClause:** <OrderBy> node.

> **rowLimit:** <RowLimit> node.

> ##### Return value
> String to be used in CAML queries

#### ViewQuery(Microsoft.SharePoint.Client.ViewScope,System.String,System.String,System.String,System.Int32)
Root <View> and <Query> nodes.
> ##### Parameters
> **scope:** View scope

> **whereClause:** <Where> node.

> **viewFields:** <ViewFields> node.

> **orderByClause:** <OrderBy> node.

> **rowLimit:** <RowLimit> node.

> ##### Return value
> String to be used in CAML queries

#### FieldValue(System.String,System.String,System.String,System.String)
Creates both a <FieldRef> and <Value> nodes combination for Where clauses.
> ##### Parameters
> **fieldName:** 

> **fieldValueType:** 

> **value:** 

> **additionalFieldRefParams:** 

> ##### Return value
> 

#### FieldValue(System.Guid,System.String,System.String,System.String)
Creates both a <FieldRef> and <Value> nodes combination for Where clauses.
> ##### Parameters
> **fieldId:** 

> **fieldValueType:** 

> **value:** 

> **additionalFieldRefParams:** 

> ##### Return value
> 

#### FieldRef(System.String)
Creates a <FieldRef> node for ViewFields clause
> ##### Parameters
> **fieldName:** 

> ##### Return value
> 

## Core.Utilities.UrlUtility
            
Static methods to modify URL paths.
        
### Methods


#### Combine(System.String,System.String[])
Combines a path and a relative path.
> ##### Parameters
> **path:** 

> **relativePaths:** 

> ##### Return value
> 

#### Combine(System.String,System.String)
Combines a path and a relative path.
> ##### Parameters
> **path:** 

> **relative:** 

> ##### Return value
> 

#### AppendQueryString(System.String,System.String)
Adds query string parameters to the end of a querystring and guarantees the proper concatenation with ? and &.
> ##### Parameters
> **path:** 

> **queryString:** 

> ##### Return value
> 

#### EnsureTrailingSlash(System.String)
Ensures that there is a trailing slash at the end of the url
> ##### Parameters
> **urlToProcess:** 

> ##### Return value
> 

## Core.Utilities.EncryptionUtility
            
Utility class that support certificate based encryption/decryption
        
### Methods


#### Encrypt(System.String,System.String)
Encrypt a piece of text based on a given certificate
> ##### Parameters
> **stringToEncrypt:** Text to encrypt

> **thumbPrint:** Thumbprint of the certificate to use

> ##### Return value
> Encrypted text

#### Decrypt(System.String,System.String)
Decrypt a piece of text based on a given certificate
> ##### Parameters
> **stringToDecrypt:** Text to decrypt

> **thumbPrint:** Thumbprint of the certificate to use

> ##### Return value
> Decrypted text

#### EncryptStringWithDPAPI(System.Security.SecureString)
Encrypts a string using the machine's DPAPI
> ##### Parameters
> **input:** String (SecureString) to encrypt

> ##### Return value
> Encrypted string

#### DecryptStringWithDPAPI(System.String)
Decrypts a DPAPI encryped string
> ##### Parameters
> **encryptedData:** Encrypted string

> ##### Return value
> Decrypted (SecureString)string

#### ToSecureString(System.String)
Converts a string to a SecureString
> ##### Parameters
> **input:** String to convert

> ##### Return value
> SecureString representation of the passed in string

#### ToInsecureString(System.Security.SecureString)
Converts a SecureString to a "regular" string
> ##### Parameters
> **input:** SecureString to convert

> ##### Return value
> A "regular" string representation of the passed SecureString

## Core.Utilities.JsonUtility
            
Utility class that supports the serialization from Json to type and vice versa
        
### Methods


#### Serialize``1(``0)
Serializes an object of type T to a json string
> ##### Parameters
> **obj:** Object to serialize

> ##### Return value
> json string

#### Deserialize``1(System.String)
Deserializes a json string to an object of type T
> ##### Parameters
> **json:** json string

> ##### Return value
> Object of type T

## Core.Utilities.SharePointContextToken
            
A JsonWebSecurityToken generated by SharePoint to authenticate to a 3rd party application and allow callbacks using a refresh token
        
### Properties

#### TargetPrincipalName
The principal name portion of the context token's "appctxsender" claim
#### RefreshToken
The context token's "refreshtoken" claim
#### CacheKey
The context token's "CacheKey" claim
#### SecurityTokenServiceUri
The context token's "SecurityTokenServiceUri" claim
#### Realm
The realm portion of the context token's "audience" claim

## Core.Utilities.MultipleSymmetricKeySecurityToken
            
Represents a security token which contains multiple security keys that are generated using symmetric algorithms.
        
### Properties

#### Id
Gets the unique identifier of the security token.
#### SecurityKeys
Gets the cryptographic keys associated with the security token.
#### ValidFrom
Gets the first instant in time at which this security token is valid.
#### ValidTo
Gets the last instant in time at which this security token is valid.
### Methods


#### Constructor
Initializes a new instance of the MultipleSymmetricKeySecurityToken class.
> ##### Parameters
> **keys:** An enumeration of Byte arrays that contain the symmetric keys.


#### Constructor
Initializes a new instance of the MultipleSymmetricKeySecurityToken class.
> ##### Parameters
> **tokenId:** The unique identifier of the security token.

> **keys:** An enumeration of Byte arrays that contain the symmetric keys.


#### MatchesKeyIdentifierClause(System.IdentityModel.Tokens.SecurityKeyIdentifierClause)
Returns a value that indicates whether the key identifier for this instance can be resolved to the specified key identifier.
> ##### Parameters
> **keyIdentifierClause:** A SecurityKeyIdentifierClause to compare to this instance

> ##### Return value
> true if keyIdentifierClause is a SecurityKeyIdentifierClause and it has the same unique identifier as the Id property; otherwise, false.

## Core.Utilities.X509CertificateUtility
            
Supporting class for certificate based operations
        
### Methods


#### LoadCertificate(System.Security.Cryptography.X509Certificates.StoreName,System.Security.Cryptography.X509Certificates.StoreLocation,System.String)
Loads a certificate from a given certificate store
> ##### Parameters
> **storeName:** Name of the certificate store

> **storeLocation:** Location of the certificate store

> **thumbprint:** Thumbprint of the certificate to load

> ##### Return value
> An certificate

#### Encrypt(System.Byte[],System.Boolean,System.Security.Cryptography.X509Certificates.X509Certificate2)
Encrypts data based on the RSACryptoServiceProvider
> ##### Parameters
> **plainData:** Bytes to encrypt

> **fOAEP:** true to perform direct System.Security.Cryptography.RSA decryption using OAEP padding

> **certificate:** Certificate to use

> ##### Return value
> Encrypted bytes

#### Decrypt(System.Byte[],System.Boolean,System.Security.Cryptography.X509Certificates.X509Certificate2)
Decrypts data based on the RSACryptoServiceProvider
> ##### Parameters
> **encryptedData:** Bytes to decrypt

> **fOAEP:** true to perform direct System.Security.Cryptography.RSA decryption using OAEP padding

> **certificate:** Certificate to use

> ##### Return value
> Decrypted bytes

#### GetPublicKey(System.Security.Cryptography.X509Certificates.X509Certificate2)
Returns the certificate public key
> ##### Parameters
> **certificate:** Certificate to operate on

> ##### Return value
> Public key of the certificate

## Core.UPAWebService.UserProfileService
            

        
### Methods


#### Constructor


#### GetUserProfileByIndex(System.Int32)


#### GetUserProfileByIndexAsync(System.Int32)


#### GetUserProfileByIndexAsync(System.Int32,System.Object)


#### CreateUserProfileByAccountName(System.String)


#### CreateUserProfileByAccountNameAsync(System.String)


#### CreateUserProfileByAccountNameAsync(System.String,System.Object)


#### GetUserProfileByName(System.String)


#### GetUserProfileByNameAsync(System.String)


#### GetUserProfileByNameAsync(System.String,System.Object)


#### GetUserProfileByGuid(System.Guid)


#### GetUserProfileByGuidAsync(System.Guid)


#### GetUserProfileByGuidAsync(System.Guid,System.Object)


#### GetUserProfileSchema


#### GetUserProfileSchemaAsync


#### GetUserProfileSchemaAsync(System.Object)


#### GetProfileSchemaNameByAccountName(System.String)


#### GetProfileSchemaNameByAccountNameAsync(System.String)


#### GetProfileSchemaNameByAccountNameAsync(System.String,System.Object)


#### GetPropertyChoiceList(System.String)


#### GetPropertyChoiceListAsync(System.String)


#### GetPropertyChoiceListAsync(System.String,System.Object)


#### ModifyUserPropertyByAccountName(System.String,OfficeDevPnP.Core.UPAWebService.PropertyData[])


#### ModifyUserPropertyByAccountNameAsync(System.String,OfficeDevPnP.Core.UPAWebService.PropertyData[])


#### ModifyUserPropertyByAccountNameAsync(System.String,OfficeDevPnP.Core.UPAWebService.PropertyData[],System.Object)


#### GetUserPropertyByAccountName(System.String,System.String)


#### GetUserPropertyByAccountNameAsync(System.String,System.String)


#### GetUserPropertyByAccountNameAsync(System.String,System.String,System.Object)


#### CreateMemberGroup(OfficeDevPnP.Core.UPAWebService.MembershipData)


#### CreateMemberGroupAsync(OfficeDevPnP.Core.UPAWebService.MembershipData)


#### CreateMemberGroupAsync(OfficeDevPnP.Core.UPAWebService.MembershipData,System.Object)


#### AddMembership(System.String,OfficeDevPnP.Core.UPAWebService.MembershipData,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### AddMembershipAsync(System.String,OfficeDevPnP.Core.UPAWebService.MembershipData,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### AddMembershipAsync(System.String,OfficeDevPnP.Core.UPAWebService.MembershipData,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Object)


#### RemoveMembership(System.String,System.Guid,System.String)


#### RemoveMembershipAsync(System.String,System.Guid,System.String)


#### RemoveMembershipAsync(System.String,System.Guid,System.String,System.Object)


#### UpdateMembershipPrivacy(System.String,System.Guid,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### UpdateMembershipPrivacyAsync(System.String,System.Guid,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### UpdateMembershipPrivacyAsync(System.String,System.Guid,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Object)


#### GetUserMemberships(System.String)


#### GetUserMembershipsAsync(System.String)


#### GetUserMembershipsAsync(System.String,System.Object)


#### GetUserOrganizations(System.String)


#### GetUserOrganizationsAsync(System.String)


#### GetUserOrganizationsAsync(System.String,System.Object)


#### GetUserColleagues(System.String)


#### GetUserColleaguesAsync(System.String)


#### GetUserColleaguesAsync(System.String,System.Object)


#### GetUserLinks(System.String)


#### GetUserLinksAsync(System.String)


#### GetUserLinksAsync(System.String,System.Object)


#### GetUserPinnedLinks(System.String)


#### GetUserPinnedLinksAsync(System.String)


#### GetUserPinnedLinksAsync(System.String,System.Object)


#### GetInCommon(System.String)


#### GetInCommonAsync(System.String)


#### GetInCommonAsync(System.String,System.Object)


#### GetCommonManager(System.String)


#### GetCommonManagerAsync(System.String)


#### GetCommonManagerAsync(System.String,System.Object)


#### GetCommonColleagues(System.String)


#### GetCommonColleaguesAsync(System.String)


#### GetCommonColleaguesAsync(System.String,System.Object)


#### GetCommonMemberships(System.String)


#### GetCommonMembershipsAsync(System.String)


#### GetCommonMembershipsAsync(System.String,System.Object)


#### AddColleague(System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Boolean)


#### AddColleagueAsync(System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Boolean)


#### AddColleagueAsync(System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Boolean,System.Object)


#### AddColleagueWithoutEmailNotification(System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Boolean)


#### AddColleagueWithoutEmailNotificationAsync(System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Boolean)


#### AddColleagueWithoutEmailNotificationAsync(System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Boolean,System.Object)


#### RemoveColleague(System.String,System.String)


#### RemoveColleagueAsync(System.String,System.String)


#### RemoveColleagueAsync(System.String,System.String,System.Object)


#### UpdateColleaguePrivacy(System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### UpdateColleaguePrivacyAsync(System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### UpdateColleaguePrivacyAsync(System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Object)


#### AddPinnedLink(System.String,System.String,System.String)


#### AddPinnedLinkAsync(System.String,System.String,System.String)


#### AddPinnedLinkAsync(System.String,System.String,System.String,System.Object)


#### UpdatePinnedLink(System.String,OfficeDevPnP.Core.UPAWebService.PinnedLinkData)


#### UpdatePinnedLinkAsync(System.String,OfficeDevPnP.Core.UPAWebService.PinnedLinkData)


#### UpdatePinnedLinkAsync(System.String,OfficeDevPnP.Core.UPAWebService.PinnedLinkData,System.Object)


#### RemovePinnedLink(System.String,System.Int32)


#### RemovePinnedLinkAsync(System.String,System.Int32)


#### RemovePinnedLinkAsync(System.String,System.Int32,System.Object)


#### AddLink(System.String,System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### AddLinkAsync(System.String,System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy)


#### AddLinkAsync(System.String,System.String,System.String,System.String,OfficeDevPnP.Core.UPAWebService.Privacy,System.Object)


#### UpdateLink(System.String,OfficeDevPnP.Core.UPAWebService.QuickLinkData)


#### UpdateLinkAsync(System.String,OfficeDevPnP.Core.UPAWebService.QuickLinkData)


#### UpdateLinkAsync(System.String,OfficeDevPnP.Core.UPAWebService.QuickLinkData,System.Object)


#### RemoveLink(System.String,System.Int32)


#### RemoveLinkAsync(System.String,System.Int32)


#### RemoveLinkAsync(System.String,System.Int32,System.Object)


#### RemoveAllLinks(System.String)


#### RemoveAllLinksAsync(System.String)


#### RemoveAllLinksAsync(System.String,System.Object)


#### RemoveAllPinnedLinks(System.String)


#### RemoveAllPinnedLinksAsync(System.String)


#### RemoveAllPinnedLinksAsync(System.String,System.Object)


#### RemoveAllColleagues(System.String)


#### RemoveAllColleaguesAsync(System.String)


#### RemoveAllColleaguesAsync(System.String,System.Object)


#### RemoveAllMemberships(System.String)


#### RemoveAllMembershipsAsync(System.String)


#### RemoveAllMembershipsAsync(System.String,System.Object)


#### GetUserProfileCount


#### GetUserProfileCountAsync


#### GetUserProfileCountAsync(System.Object)


#### AddSuggestions(OfficeDevPnP.Core.UPAWebService.SuggestionType,System.String[],System.Double[])


#### AddSuggestionsAsync(OfficeDevPnP.Core.UPAWebService.SuggestionType,System.String[],System.Double[])


#### AddSuggestionsAsync(OfficeDevPnP.Core.UPAWebService.SuggestionType,System.String[],System.Double[],System.Object)


#### GetProfileSchemaNames


#### GetProfileSchemaNamesAsync


#### GetProfileSchemaNamesAsync(System.Object)


#### GetProfileSchema(System.String)


#### GetProfileSchemaAsync(System.String)


#### GetProfileSchemaAsync(System.String,System.Object)


#### GetLeaders


#### GetLeadersAsync


#### GetLeadersAsync(System.Object)


#### AddLeader(System.String)


#### AddLeaderAsync(System.String)


#### AddLeaderAsync(System.String,System.Object)


#### RemoveLeader(System.String)


#### RemoveLeaderAsync(System.String)


#### RemoveLeaderAsync(System.String,System.Object)


#### CancelAsync(System.Object)


## Core.UPAWebService.GetUserProfileByIndexResult
            

        
### Properties

#### NextValue

#### UserProfile

#### Colleagues

#### QuickLinks

#### PinnedLinks

#### Memberships


## Core.UPAWebService.PropertyData
            

        
### Properties

#### IsPrivacyChanged

#### IsValueChanged

#### Name

#### Privacy

#### Values


## Core.UPAWebService.Privacy
            

        
### Fields

#### Public

#### Contacts

#### Organization

#### Manager

#### Private

#### NotSet


## Core.UPAWebService.ValueData
            

        
### Properties

#### Value


## Core.UPAWebService.Leader
            

        
### Properties

#### AccountName

#### Valid

#### ManagerAccountName

#### ReportCount


## Core.UPAWebService.InCommonData
            

        
### Properties

#### Manager

#### Colleagues

#### Memberships


## Core.UPAWebService.ContactData
            

        
### Properties

#### AccountName

#### Privacy

#### Name

#### IsInWorkGroup

#### Group

#### Email

#### Title

#### Url

#### UserProfileID

#### ID


## Core.UPAWebService.MembershipData
            

        
### Properties

#### Source

#### MemberGroup

#### Group

#### DisplayName

#### Privacy

#### MailNickname

#### Url

#### ID

#### MemberGroupID


## Core.UPAWebService.MembershipSource
            

        
### Fields

#### DistributionList

#### SharePointSite

#### Other


## Core.UPAWebService.MemberGroupData
            

        
### Properties

#### SourceInternal

#### SourceReference


## Core.UPAWebService.OrganizationProfileData
            

        
### Properties

#### DisplayName

#### RecordID


## Core.UPAWebService.PropertyInfo
            

        
### Properties

#### Name

#### Description

#### DisplayOrder

#### MaximumShown

#### IsAdminEditable

#### IsSearchable

#### IsSystem

#### ManagedPropertyName

#### DisplayName

#### Type

#### AllowPolicyOverride

#### DefaultPrivacy

#### IsAlias

#### IsColleagueEventLog

#### IsRequired

#### IsUserEditable

#### IsVisibleOnEditor

#### IsVisibleOnViewer

#### IsReplicable

#### UserOverridePrivacy

#### Length

#### IsImported

#### IsMultiValue

#### ChoiceType

#### TermSetId


## Core.UPAWebService.ChoiceTypes
            

        
### Fields

#### Off

#### None

#### Open

#### Closed


## Core.UPAWebService.SPTimeZone
            

        
### Properties

#### ID


## Core.UPAWebService.PinnedLinkData
            

        
### Properties

#### Name

#### Url

#### ID


## Core.UPAWebService.QuickLinkData
            

        
### Properties

#### Name

#### Group

#### Privacy

#### Url

#### ID


## Core.UPAWebService.SuggestionType
            

        
### Fields

#### Colleague

#### Keyword


## Core.UPAWebService.GetUserProfileByIndexCompletedEventHandler
            

        

## Core.UPAWebService.GetUserProfileByIndexCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.CreateUserProfileByAccountNameCompletedEventHandler
            

        

## Core.UPAWebService.CreateUserProfileByAccountNameCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetUserProfileByNameCompletedEventHandler
            

        

## Core.UPAWebService.GetUserProfileByNameCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetUserProfileByGuidCompletedEventHandler
            

        

## Core.UPAWebService.GetUserProfileByGuidCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetUserProfileSchemaCompletedEventHandler
            

        

## Core.UPAWebService.GetUserProfileSchemaCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetProfileSchemaNameByAccountNameCompletedEventHandler
            

        

## Core.UPAWebService.GetProfileSchemaNameByAccountNameCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetPropertyChoiceListCompletedEventHandler
            

        

## Core.UPAWebService.GetPropertyChoiceListCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.ModifyUserPropertyByAccountNameCompletedEventHandler
            

        

## Core.UPAWebService.GetUserPropertyByAccountNameCompletedEventHandler
            

        

## Core.UPAWebService.GetUserPropertyByAccountNameCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.CreateMemberGroupCompletedEventHandler
            

        

## Core.UPAWebService.AddMembershipCompletedEventHandler
            

        

## Core.UPAWebService.AddMembershipCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.RemoveMembershipCompletedEventHandler
            

        

## Core.UPAWebService.UpdateMembershipPrivacyCompletedEventHandler
            

        

## Core.UPAWebService.GetUserMembershipsCompletedEventHandler
            

        

## Core.UPAWebService.GetUserMembershipsCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetUserOrganizationsCompletedEventHandler
            

        

## Core.UPAWebService.GetUserOrganizationsCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetUserColleaguesCompletedEventHandler
            

        

## Core.UPAWebService.GetUserColleaguesCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetUserLinksCompletedEventHandler
            

        

## Core.UPAWebService.GetUserLinksCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetUserPinnedLinksCompletedEventHandler
            

        

## Core.UPAWebService.GetUserPinnedLinksCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetInCommonCompletedEventHandler
            

        

## Core.UPAWebService.GetInCommonCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetCommonManagerCompletedEventHandler
            

        

## Core.UPAWebService.GetCommonManagerCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetCommonColleaguesCompletedEventHandler
            

        

## Core.UPAWebService.GetCommonColleaguesCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetCommonMembershipsCompletedEventHandler
            

        

## Core.UPAWebService.GetCommonMembershipsCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.AddColleagueCompletedEventHandler
            

        

## Core.UPAWebService.AddColleagueCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.AddColleagueWithoutEmailNotificationCompletedEventHandler
            

        

## Core.UPAWebService.AddColleagueWithoutEmailNotificationCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.RemoveColleagueCompletedEventHandler
            

        

## Core.UPAWebService.UpdateColleaguePrivacyCompletedEventHandler
            

        

## Core.UPAWebService.AddPinnedLinkCompletedEventHandler
            

        

## Core.UPAWebService.AddPinnedLinkCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.UpdatePinnedLinkCompletedEventHandler
            

        

## Core.UPAWebService.RemovePinnedLinkCompletedEventHandler
            

        

## Core.UPAWebService.AddLinkCompletedEventHandler
            

        

## Core.UPAWebService.AddLinkCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.UpdateLinkCompletedEventHandler
            

        

## Core.UPAWebService.RemoveLinkCompletedEventHandler
            

        

## Core.UPAWebService.RemoveAllLinksCompletedEventHandler
            

        

## Core.UPAWebService.RemoveAllPinnedLinksCompletedEventHandler
            

        

## Core.UPAWebService.RemoveAllColleaguesCompletedEventHandler
            

        

## Core.UPAWebService.RemoveAllMembershipsCompletedEventHandler
            

        

## Core.UPAWebService.GetUserProfileCountCompletedEventHandler
            

        

## Core.UPAWebService.GetUserProfileCountCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.AddSuggestionsCompletedEventHandler
            

        

## Core.UPAWebService.GetProfileSchemaNamesCompletedEventHandler
            

        

## Core.UPAWebService.GetProfileSchemaNamesCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetProfileSchemaCompletedEventHandler
            

        

## Core.UPAWebService.GetProfileSchemaCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.GetLeadersCompletedEventHandler
            

        

## Core.UPAWebService.GetLeadersCompletedEventArgs
            

        
### Properties

#### Result


## Core.UPAWebService.AddLeaderCompletedEventHandler
            

        

## Core.UPAWebService.RemoveLeaderCompletedEventHandler
            

        

## Core.WebAPI.WebAPIContexCacheItem
            
This class holds the information that's being cached as part of the WebAPI service implementation
        
### Properties

#### AccessToken
The SharePoint Access token
#### RefreshToken
The SharePoint Refresh token
#### SharePointServiceContext
The information initially used to register the SharePoint app to the WebAPI service

## Core.WebAPI.WebAPIContext
            
This class holds the information that's passed from the SharePoint app to the "Register" WebAPI service call
        
### Properties

#### CacheKey
The cacheKey that will be used. The cache key is unique for each combination of user name, user name issuer, application, and farm.
#### Token
The SharePoint context token. This will be used at the WebAPI level to obtain an access token
#### HostWebUrl
Url of the SharePoint host web. Needed to obtain an access token
#### AppWebUrl
Url if the SharePoint app web. Needed to obtain an access token
#### ClientId
ClientId of the SharePoint app that's being registered. Needed to obtain an access token
#### ClientSecret
ClientSecret of the SharePoint app that's being registered. Needed to obtain an access token
#### HostedAppHostName
If the AppWebUrl is null then this value will be used. Needed to obtain an access token
#### 
Singleton instance to access this class
### Methods


#### 
Adds an item to the cache. Updates if the item already existed
> ##### Parameters
> **cacheKey:** Key to cache the item

> **sharePointServiceContextCacheItem:** A object


#### 
Gets an item from the cache
> ##### Parameters
> **cacheKey:** Key to retrieve an item from cache

> ##### Return value
> A object

## Core.WebAPI.WebAPIContextCache
            
Simple cache implementation based on the singleton pattern. Caches the SharePoint access token, refresh token and the information passed during service "registration". All of this information is wrapped in a object.
        
### Properties

#### Instance
Singleton instance to access this class
### Methods


#### Put(System.String,OfficeDevPnP.Core.WebAPI.WebAPIContexCacheItem)
Adds an item to the cache. Updates if the item already existed
> ##### Parameters
> **cacheKey:** Key to cache the item

> **sharePointServiceContextCacheItem:** A object


#### Get(System.String)
Gets an item from the cache
> ##### Parameters
> **cacheKey:** Key to retrieve an item from cache

> ##### Return value
> A object

## Core.WebAPI.WebAPIHelper
            
This class provides helper methods that can be used to protect WebAPI services and to provide a way to reinstantiate a contextobject in the service call.
        
### Fields

#### SERVICES_TOKEN
This is the name of the cookie that will hold the cachekey.
### Methods


#### HasCacheEntry(System.Web.Http.Controllers.HttpControllerContext)
Checks if this request has a servicesToken cookie. To be used from inside the WebAPI.
> ##### Parameters
> **httpControllerContext:** Information about the HTTP request that reached the WebAPI controller

> ##### Return value
> True if cookie is available and not empty, false otherwise

#### GetClientContext(System.Web.Http.Controllers.HttpControllerContext)
Creates a ClientContext token for the incoming WebAPI request. This is done by - looking up the servicesToken - extracting the cacheKey - get the AccessToken from cache. If the AccessToken is expired a new one is requested using the refresh token - creation of a ClientContext object based on the AccessToken
> ##### Parameters
> **httpControllerContext:** Information about the HTTP request that reached the WebAPI controller

> ##### Return value
> A valid ClientContext object

#### AddToCache(OfficeDevPnP.Core.WebAPI.WebAPIContext)
Uses the information regarding the requesting app to obtain an access token and caches that using the cachekey. This method is called from the Register WebAPI service api.
> ##### Parameters
> **sharePointServiceContext:** Object holding information about the requesting SharePoint app


#### RegisterWebAPIService(System.Web.UI.Page,System.String,System.Uri)
This method needs to be called from a code behind of the SharePoint app startup page (default.aspx). It registers the calling SharePoint app by calling a specific "Register" api in your WebAPI service. Note: Given that method is async you'll need to add the Async="true" page directive to the page that uses this method.
> ##### Parameters
> **page:** The page object, needed to insert the services token cookie and read the querystring

> **apiRequest:** Route to the "Register" API

> **serviceEndPoint:** Optional Uri to the WebAPI service endpoint. If null then the assumption is taken that the WebAPI is hosted together with the page making this call


## EnumerationExtensions
            
Extension methods to make working with Enum values easier. Copied from http://hugoware.net/blog/enumeration-extensions-2-0.
        
### Methods


#### Include``1(System.Enum,``0)
Includes an enumerated type and returns the new value

#### Remove``1(System.Enum,``0)
Removes an enumerated type and returns the new value

#### Has``1(System.Enum,``0)
Checks if an enumerated type contains a value

#### Missing``1(System.Enum,``0)
Checks if an enumerated type is missing a value

## SafeConvertExtensions
            
Safely convert strings to specified types.
        
### Methods


#### ToBoolean(System.String,System.Boolean)
Converts the input string to a boolean and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.

> **defaultValue:** A default value to return for a null input value.


#### ToBoolean(System.String)
Converts the input string to a boolean and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.


#### ToInt32(System.String,System.Int32)
Converts the input string to a Int32 and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.

> **defaultValue:** A default value to return for a null input value.


#### ToInt32(System.String)
Converts the input string to a Int64 and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.


#### ToInt64(System.String,System.Int32)
Converts the input string to a Int32 and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.

> **defaultValue:** A default value to return for a null input value.


#### ToInt64(System.String)
Converts the input string to a Int32 and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.


#### ToDouble(System.String,System.Double)
Converts the input string to a double and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.

> **defaultValue:** A default value to return for a null input value.


#### ToDouble(System.String)
Converts the input string to a double and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.


#### ToGuid(System.String)
Converts the input string to a Guid and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.


#### ToGuid(System.String,System.Guid)
Converts the input string to a Guid and if null, it returns the default value.
> ##### Parameters
> **input:** Input string.

> **defaultValue:** A default value to return for a null input value.
