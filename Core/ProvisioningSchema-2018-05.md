
# PnP Provisioning Schema
----------
* Topic automatically generated on 5/3/2018*

## Namespace
The namespace of the PnP Provisioning Schema is:

http://schemas.dev.office.com/PnP/2018/05/ProvisioningSchema

All the elements have to be declared with that namespace reference.

## Root Elements
Here follows the list of root elements available in the PnP Provisioning Schema.
  
<a name="provisioning"></a>
### Provisioning


```xml
<pnp:Provisioning
      xmlns:pnp="http://schemas.dev.office.com/PnP/2018/05/ProvisioningSchema">
   <pnp:Preferences />
   <pnp:Localizations />
   <pnp:Tenant />
   <pnp:Templates />
   <pnp:Sequence />
   <pnp:ImportSequence />
</pnp:Provisioning>
```


Here follow the available child elements for the Provisioning element.


Element|Type|Description
-------|----|-----------
Preferences|[Preferences](#preferences)|The mandatory section of preferences for the current provisioning definition.
Localizations|[Localizations](#localizations)|An optional list of localizations files to include.
Tenant|[Tenant](#tenant)|Entry point to manage tenant-wide settings.
Templates|[Templates](#templates)|An optional section made of provisioning templates.
Sequence|[Sequence](#sequence)|An optional section made of provisioning sequences, which can include Sites, Site Collections, Taxonomies, Provisioning Templates, etc.
ImportSequence|[ImportSequence](#importsequence)|Imports sequences from an external file. All current properties should be sent to that file.
<a name="provisioningtemplate"></a>
### ProvisioningTemplate
Represents the root element of the SharePoint Provisioning Template.

```xml
<pnp:ProvisioningTemplate
      xmlns:pnp="http://schemas.dev.office.com/PnP/2018/05/ProvisioningSchema"
      ID="xsd:ID"
      Version="xsd:decimal"
      BaseSiteTemplate="xsd:string"
      ImagePreviewUrl="xsd:string"
      DisplayName="xsd:string"
      Description="xsd:string"
      TemplateCultureInfo="pnp:ReplaceableString"
      Scope="">
   <pnp:Properties />
   <pnp:SitePolicy />
   <pnp:WebSettings />
   <pnp:RegionalSettings />
   <pnp:SupportedUILanguages />
   <pnp:AuditSettings />
   <pnp:PropertyBagEntries />
   <pnp:Security />
   <pnp:Navigation />
   <pnp:SiteFields />
   <pnp:ContentTypes />
   <pnp:Lists />
   <pnp:Features />
   <pnp:CustomActions />
   <pnp:Files />
   <pnp:Pages />
   <pnp:TermGroups />
   <pnp:ComposedLook />
   <pnp:Workflows />
   <pnp:SearchSettings />
   <pnp:Publishing />
   <pnp:ApplicationLifecycleManagement />
   <pnp:Providers />
   <pnp:SiteWebhooks />
   <pnp:ClientSidePages />
</pnp:ProvisioningTemplate>
```


Here follow the available child elements for the ProvisioningTemplate element.


Element|Type|Description
-------|----|-----------
Properties|[ProvisioningTemplateProperties](#provisioningtemplateproperties)|A set of custom Properties for the Provisioning Template, optional element.
SitePolicy|[ReplaceableString](#replaceablestring)|The Site Policy of the Provisioning Template, optional element.
WebSettings|[WebSettings](#websettings)|Section of Settings for the current Web Site, optional element.
RegionalSettings|[RegionalSettings](#regionalsettings)|The Regional Settings of the Provisioning Template, optional element.
SupportedUILanguages|[SupportedUILanguages](#supporteduilanguages)|The Supported UI Languages for the Provisioning Template, optional element.
AuditSettings|[AuditSettings](#auditsettings)|The Audit Settings for the Provisioning Template, optional element.
PropertyBagEntries|[PropertyBagEntries](#propertybagentries)|The Property Bag entries of the Provisioning Template, optional collection of elements.
Security|[Security](#security)|The Security configurations of the Provisioning Template, optional collection of elements.
Navigation|[Navigation](#navigation)|The Navigation configurations of the Provisioning Template, optional collection of elements.
SiteFields|[SiteFields](#sitefields)|The Site Columns of the Provisioning Template, optional element.
ContentTypes|[ContentTypes](#contenttypes)|The Content Types of the Provisioning Template, optional element.
Lists|[Lists](#lists)|The Lists instances of the Provisioning Template, optional element.
Features|[Features](#features)|The Features (Site or Web) to activate or deactivate while applying the Provisioning Template, optional collection of elements.
CustomActions|[CustomActions](#customactions)|The Custom Actions (Site or Web) to provision with the Provisioning Template, optional element.
Files|[Files](#files)|The Files to provision into the target Site through the Provisioning Template, optional element.
Pages|[Pages](#pages)|The Pages to provision into the target Site through the Provisioning Template, optional element.
TermGroups|[TermGroups](#termgroups)|The TermGroups element allows provisioning one or more TermGroups into the target Site, optional element.
ComposedLook|[ComposedLook](#composedlook)|The ComposedLook for the Provisioning Template, optional element.
Workflows|[Workflows](#workflows)|The Workflows for the Provisioning Template, optional element.
SearchSettings|[SearchSettings](#searchsettings)|The Search Settings for the Provisioning Template, optional element.
Publishing|[Publishing](#publishing)|The Publishing capabilities configuration for the Provisioning Template, optional element.
ApplicationLifecycleManagement|[ApplicationLifecycleManagement](#applicationlifecyclemanagement)|Entry point to manage the ALM of SharePoint Add-Ins and SharePoint Framework solutions at the site collection level.
Providers|[Providers](#providers)|The Extensiblity Providers to invoke while applying the Provisioning Template, optional collection of elements.
SiteWebhooks|[SiteWebhooksList](#sitewebhookslist)|Defines any list of Webhooks for the current site.
ClientSidePages|[ClientSidePages](#clientsidepages)|The Client-side Pages to provision into the target Site through the Provisioning Template, optional element.

Here follow the available attributes for the ProvisioningTemplate element.


Attibute|Type|Description
--------|----|-----------
ID|xsd:ID|The ID of the Provisioning Template, required attribute.
Version|xsd:decimal|The Version of the Provisioning Template, optional attribute.
BaseSiteTemplate|xsd:string|The Base SiteTemplate of the Provisioning Template, optional attribute.
ImagePreviewUrl|xsd:string|The Image Preview Url of the Provisioning Template, optional attribute.
DisplayName|xsd:string|The Display Name of the Provisioning Template, optional attribute.
Description|xsd:string|The Description of the Provisioning Template, optional attribute.
TemplateCultureInfo|ReplaceableString|The default CultureInfo of the Provisioning Template, used to format all input values, optional attribute.
Scope||Declares the target scope of the current Provisioning Template


## Child Elements and Complex Types
Here follows the list of all the other child elements and complex types that can be used in the PnP Provisioning Schema.
<a name="preferences"></a>
### Preferences
General settings for a Provisioning file.

```xml
<pnp:Preferences
      Version="xsd:string"
      Author="xsd:string"
      Generator="xsd:string">
   <pnp:Parameters />
</pnp:Preferences>
```


Here follow the available child elements for the Preferences element.


Element|Type|Description
-------|----|-----------
Parameters|[Parameters](#parameters)|Definition of parameters that can be used as replacement within templates and provisioning objects.

Here follow the available attributes for the Preferences element.


Attibute|Type|Description
--------|----|-----------
Version|xsd:string|Provisioning File Version number, optional attribute.
Author|xsd:string|Provisioning File Author name, optional attribute.
Generator|xsd:string|Name of the tool generating this Provisioning File, optional attribute.
<a name="parameters"></a>
### Parameters
Definition of parameters that can be used as replacement within templates and provisioning objects.

```xml
<pnp:Parameters>
   <pnp:Parameter />
</pnp:Parameters>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Parameter|[Parameter](#parameter)|A Parameter that can be used as a replacement within templates and provisioning objects.
<a name="localizations"></a>
### Localizations
An optional list of localizations files to include.

```xml
<pnp:Localizations>
   <pnp:Localization />
</pnp:Localizations>
```


Here follow the available child elements for the Localizations element.


Element|Type|Description
-------|----|-----------
Localization|[Localization](#localization)|A Localization element
<a name="localization"></a>
### Localization
A Localization element

```xml
<pnp:Localization
      LCID="xsd:int"
      Name="xsd:string"
      ResourceFile="xsd:string">
</pnp:Localization>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
LCID|xsd:int|The Locale ID of a Localization Language, required attribute.
Name|xsd:string|The Name of a Localization Language, required attribute.
ResourceFile|xsd:string|The path to the .RESX (XML) resource file for the current Localization, required attribute.
<a name="tenant"></a>
### Tenant
Element to manage tenant-wide settings.

```xml
<pnp:Tenant>
   <pnp:AppCatalog />
   <pnp:WebApiPermissions />
   <pnp:ContentDeliveryNetwork />
   <pnp:SiteDesigns />
   <pnp:SiteScripts />
   <pnp:StorageEntities />
</pnp:Tenant>
```


Here follow the available child elements for the Tenant element.


Element|Type|Description
-------|----|-----------
AppCatalog|[AppCatalog](#appcatalog)|Entry point for the tenant-wide AppCatalog
WebApiPermissions|[WebApiPermissions](#webapipermissions)|Entry point for the tenant-wide Web API permissions
ContentDeliveryNetwork|[ContentDeliveryNetwork](#contentdeliverynetwork)|Entry point for the tenant-wide Content Delivery Network
SiteDesigns|[SiteDesigns](#sitedesigns)|Entry point for the tenant-wide Site Designs
SiteScripts|[SiteScripts](#sitescripts)|Entry point for the tenant-wide Site Scripts
StorageEntities|[StorageEntities](#storageentities)|Entry point for the tenant-wide properties (Storage Entities)
<a name="webapipermissions"></a>
### WebApiPermissions
Collection of tenant-wide Web API permissions.

```xml
<pnp:WebApiPermissions>
   <pnp:WebApiPermission />
</pnp:WebApiPermissions>
```


Here follow the available child elements for the WebApiPermissions element.


Element|Type|Description
-------|----|-----------
WebApiPermission|[WebApiPermission](#webapipermission)|
<a name="webapipermission"></a>
### WebApiPermission
A single tenant-wide Web API permission.

```xml
<pnp:WebApiPermission
      Resource="xsd:string"
      Scope="xsd:string">
</pnp:WebApiPermission>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Resource|xsd:string|The target resource for a Web API permission, required attribute.
Scope|xsd:string|The target resource for a Web API permission, required attribute.
<a name="sitedesigns"></a>
### SiteDesigns
Collection of tenant-wide Site Designs

```xml
<pnp:SiteDesigns>
   <pnp:SiteDesign />
</pnp:SiteDesigns>
```


Here follow the available child elements for the SiteDesigns element.


Element|Type|Description
-------|----|-----------
SiteDesign|[SiteDesign](#sitedesign)|Defines a single tenant-wide Site Design
<a name="sitedesign"></a>
### SiteDesign
Defines a single tenant-wide Site Design

```xml
<pnp:SiteDesign
      Title="xsd:string"
      Description="xsd:string"
      IsDefault="xsd:boolean"
      WebTemplate=""
      PreviewImageUrl="xsd:string"
      PreviewImageAltText="xsd:string"
      Overwrite="xsd:boolean">
   <pnp:SiteScripts />
   <pnp:Grants />
</pnp:SiteDesign>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
SiteScripts|[SiteScripts](#sitescripts)|A collection of Site Scripts references for the current Site Design
Grants|[Grants](#grants)|A collection of Grants for the Site Design

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Title|xsd:string|The Title of the Site Design
Description|xsd:string|The Description of the Site Design, optional attribute.
IsDefault|xsd:boolean|Defines whether the Site Design is default or not, optional attribute (default: false).
WebTemplate||Defines whether the Site Design is default or not, required attribute.
PreviewImageUrl|xsd:string|The URL of the preview image for the Site Design, optional attribute.
PreviewImageAltText|xsd:string|The alternate text for the preview image of the Site Design, optional attribute.
Overwrite|xsd:boolean|Defines whether the Site Design should be overwritten in case of existence, optional attribute (default:true).
<a name="sitescripts"></a>
### SiteScripts
Collection of tenant-wide Site Scripts

```xml
<pnp:SiteScripts>
   <pnp:SiteScript />
</pnp:SiteScripts>
```


Here follow the available child elements for the SiteScripts element.


Element|Type|Description
-------|----|-----------
SiteScript|[SiteScript](#sitescript)|Defines a single tenant-wide Site Script
<a name="sitescript"></a>
### SiteScript
Defines a single tenant-wide Site Script

```xml
<pnp:SiteScript
      Title="xsd:string"
      Description="xsd:string"
      JsonFilePath="xsd:string"
      Overwrite="xsd:boolean">
</pnp:SiteScript>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Title|xsd:string|The Title of the Site Script, required attribute.
Description|xsd:string|The Description of the Site Script, required attribute.
JsonFilePath|xsd:string|The path of the JSON file defining the Site Script, required attribute.
Overwrite|xsd:boolean|Defines whether the Site Script should be overwritten in case of existence, optional attribute (default:true).
<a name="storageentities"></a>
### StorageEntities
Collection of tenant-wide properties (Storage Entities)

```xml
<pnp:StorageEntities
      Comment="xsd:string"
      Description="xsd:string">
   <pnp:StorageEntity />
</pnp:StorageEntities>
```


Here follow the available child elements for the StorageEntities element.


Element|Type|Description
-------|----|-----------
StorageEntity|[StorageEntity](#storageentity)|Defines a single tenant-wide property (Storage Entity)

Here follow the available attributes for the StorageEntities element.


Attibute|Type|Description
--------|----|-----------
Comment|xsd:string|The Comment of the tenant-wide property, optional attribute.
Description|xsd:string|The Description of the tenant-wide property, optional attribute.
<a name="storageentity"></a>
### StorageEntity
Defines a single tenant-wide property (Storage Entity)

```xml
<pnp:StorageEntity
      Comment="xsd:string"
      Description="xsd:string">
</pnp:StorageEntity>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Comment|xsd:string|The Comment of the tenant-wide property, optional attribute.
Description|xsd:string|The Description of the tenant-wide property, optional attribute.
<a name="templates"></a>
### Templates
SharePoint Templates, which can be inline or references to external files.

```xml
<pnp:Templates
      ID="xsd:ID">
   <pnp:ProvisioningTemplateFile />
   <pnp:ProvisioningTemplateReference />
   <pnp:ProvisioningTemplate />
</pnp:Templates>
```


Here follow the available child elements for the Templates element.


Element|Type|Description
-------|----|-----------
ProvisioningTemplateFile|[ProvisioningTemplateFile](#provisioningtemplatefile)|Reference to an external template file, which will be based on the current schema but will focus only on the SharePointProvisioningTemplate section.
ProvisioningTemplateReference|[ProvisioningTemplateReference](#provisioningtemplatereference)|Reference to another template by ID.
ProvisioningTemplate|[ProvisioningTemplate](#provisioningtemplate)|

Here follow the available attributes for the Templates element.


Attibute|Type|Description
--------|----|-----------
ID|xsd:ID|A unique identifier of the Templates collection, optional attribute.
<a name="sitefields"></a>
### SiteFields
The Site Columns of the Provisioning Template, optional element.

```xml
<pnp:SiteFields>
   <!-- Any other XML content -->
</pnp:SiteFields>
```

<a name="contenttypes"></a>
### ContentTypes
The Content Types of the Provisioning Template, optional element.

```xml
<pnp:ContentTypes>
   <pnp:ContentType />
</pnp:ContentTypes>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
ContentType|[ContentType](#contenttype)|
<a name="lists"></a>
### Lists
The Lists instances of the Provisioning Template, optional element.

```xml
<pnp:Lists>
   <pnp:ListInstance />
</pnp:Lists>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
ListInstance|[ListInstance](#listinstance)|
<a name="files"></a>
### Files
The Files to provision into the target Site through the Provisioning Template, optional element.

```xml
<pnp:Files>
   <pnp:File />
   <pnp:Directory />
</pnp:Files>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
File|[File](#file)|
Directory|[Directory](#directory)|
<a name="termgroups"></a>
### TermGroups
The TermGroups element allows provisioning one or more TermGroups into the target Site, optional element.

```xml
<pnp:TermGroups>
   <pnp:TermGroup />
</pnp:TermGroups>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
TermGroup|[TermGroup](#termgroup)|The TermGroup element to provision into the target Site through the Provisioning Template, optional element.
<a name="searchsettings"></a>
### SearchSettings
The Search Settings for the Provisioning Template, optional element.

```xml
<pnp:SearchSettings>
   <pnp:SiteSearchSettings />
   <pnp:WebSearchSettings />
</pnp:SearchSettings>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
SiteSearchSettings|[SiteSearchSettings](#sitesearchsettings)|The Search Settings for the Site Collection, optional element.
WebSearchSettings|[WebSearchSettings](#websearchsettings)|The Search Settings for the Site, optional element.
<a name="providers"></a>
### Providers
The Extensiblity Providers to invoke while applying the Provisioning Template, optional collection of elements.

```xml
<pnp:Providers>
   <pnp:Provider />
</pnp:Providers>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Provider|[Provider](#provider)|
<a name="provisioningtemplateproperties"></a>
### ProvisioningTemplateProperties
A set of custom Properties for the Provisioning Template.

```xml
<pnp:ProvisioningTemplateProperties>
   <pnp:Property />
</pnp:ProvisioningTemplateProperties>
```


Here follow the available child elements for the ProvisioningTemplateProperties element.


Element|Type|Description
-------|----|-----------
Property|[StringDictionaryItem](#stringdictionaryitem)|A custom Property for the Provisioning Template, collection of elements.
<a name="websettings"></a>
### WebSettings
Section of Settings for the current Web Site, optional element.

```xml
<pnp:WebSettings
      RequestAccessEmail="xsd:string"
      NoCrawl="xsd:boolean"
      WelcomePage="xsd:string"
      Title="xsd:string"
      Description="xsd:string"
      SiteLogo="xsd:string"
      AlternateCSS="xsd:string"
      MasterPageUrl="xsd:string"
      CustomMasterPageUrl="xsd:string"
      HubSiteUrl="xsd:string"
      CommentsOnSitePagesDisabled="xsd:boolean">
</pnp:WebSettings>
```


Here follow the available attributes for the WebSettings element.


Attibute|Type|Description
--------|----|-----------
RequestAccessEmail|xsd:string|The email address to which any access request will be sent, optional attribute.
NoCrawl|xsd:boolean|Defines whether the site has to be crawled or not, optional attribute.
WelcomePage|xsd:string|Defines the Welcome Page (Home Page) of the site to which the Provisioning Template is applied, optional attribute. The page does not necessarily need to be in the current template, can be an already existing one.
Title|xsd:string|The Title of the Site, optional attribute.
Description|xsd:string|The Description of the Site, optional attribute.
SiteLogo|xsd:string|The SiteLogo of the Site, optional attribute.
AlternateCSS|xsd:string|The AlternateCSS of the Site, optional attribute.
MasterPageUrl|xsd:string|The MasterPage URL of the Site, optional attribute.
CustomMasterPageUrl|xsd:string|The Custom MasterPage URL of the Site, optional attribute.
HubSiteUrl|xsd:string|The URL of the Hub Site to associate the site to, optional attribute. If it is empty, you disassociate it from any Hub Site.
CommentsOnSitePagesDisabled|xsd:boolean|Enables or disables comments on client side pages.
<a name="regionalsettings"></a>
### RegionalSettings
Defines the Regional Settings for a site.

```xml
<pnp:RegionalSettings
      AdjustHijriDays="xsd:int"
      AlternateCalendarType="pnp:CalendarType"
      CalendarType="pnp:CalendarType"
      Collation="xsd:int"
      FirstDayOfWeek="pnp:DayOfWeek"
      FirstWeekOfYear="xsd:int"
      LocaleId="xsd:int"
      ShowWeeks="xsd:boolean"
      Time24="xsd:boolean"
      TimeZone="pnp:ReplaceableInt"
      WorkDayEndHour="pnp:WorkHour"
      WorkDays="xsd:int"
      WorkDayStartHour="pnp:WorkHour">
</pnp:RegionalSettings>
```


Here follow the available attributes for the RegionalSettings element.


Attibute|Type|Description
--------|----|-----------
AdjustHijriDays|xsd:int|The number of days to extend or reduce the current month in Hijri calendars, optional attribute.
AlternateCalendarType|CalendarType|The Alternate Calendar type that is used on the server, optional attribute.
CalendarType|CalendarType|The Calendar Type that is used on the server, optional attribute.
Collation|xsd:int|The Collation that is used on the site, optional attribute.
FirstDayOfWeek|DayOfWeek|The First Day of the Week used in calendars on the server, optional attribute.
FirstWeekOfYear|xsd:int|The First Week of the Year used in calendars on the server, optional attribute.
LocaleId|xsd:int|The Locale Identifier in use on the server, optional attribute.
ShowWeeks|xsd:boolean|Defines whether to display the week number in day or week views of a calendar, optional attribute.
Time24|xsd:boolean|Defines whether to use a 24-hour time format in representing the hours of the day, optional attribute.
TimeZone|ReplaceableInt|The Time Zone that is used on the server, optional attribute.
WorkDayEndHour|WorkHour|The the default hour at which the work day ends on the calendar that is in use on the server, optional attribute.
WorkDays|xsd:int|The work days of Web site calendars, optional attribute.
WorkDayStartHour|WorkHour|The the default hour at which the work day starts on the calendar that is in use on the server, optional attribute.
<a name="supporteduilanguages"></a>
### SupportedUILanguages
Defines the Supported UI Languages for a site.

```xml
<pnp:SupportedUILanguages>
   <pnp:SupportedUILanguage />
</pnp:SupportedUILanguages>
```


Here follow the available child elements for the SupportedUILanguages element.


Element|Type|Description
-------|----|-----------
SupportedUILanguage|[SupportedUILanguage](#supporteduilanguage)|Defines a single Supported UI Language for a site.
<a name="supporteduilanguage"></a>
### SupportedUILanguage
Defines a single Supported UI Language for a site.

```xml
<pnp:SupportedUILanguage
      LCID="xsd:int">
</pnp:SupportedUILanguage>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
LCID|xsd:int|The Locale ID of a Supported UI Language, required attribute.
<a name="auditsettings"></a>
### AuditSettings
The Audit Settings for the Provisioning Template, optional collection of elements.

```xml
<pnp:AuditSettings
      AuditLogTrimmingRetention="xsd:int"
      TrimAuditLog="xsd:boolean">
   <pnp:Audit />
</pnp:AuditSettings>
```


Here follow the available child elements for the AuditSettings element.


Element|Type|Description
-------|----|-----------
Audit|[Audit](#audit)|A single Audit setting defined by an AuditFlag.

Here follow the available attributes for the AuditSettings element.


Attibute|Type|Description
--------|----|-----------
AuditLogTrimmingRetention|xsd:int|The Audit Log Trimming Retention for Audits, optional attribute.
TrimAuditLog|xsd:boolean|A flag to enable Audit Log Trimming, optional attribute.
<a name="audit"></a>
### Audit
A single Audit setting defined by an AuditFlag.

```xml
<pnp:Audit
      AuditFlag="">
</pnp:Audit>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
AuditFlag||An Audit Flag for a single Audit setting, required attribute.
<a name="propertybagentries"></a>
### PropertyBagEntries
The Property Bag entries of the Provisioning Template, optional collection of elements.

```xml
<pnp:PropertyBagEntries>
   <pnp:PropertyBagEntry />
</pnp:PropertyBagEntries>
```


Here follow the available child elements for the PropertyBagEntries element.


Element|Type|Description
-------|----|-----------
PropertyBagEntry|[PropertyBagEntry](#propertybagentry)|
<a name="security"></a>
### Security
The Security configurations of the Provisioning Template, optional collection of elements.

```xml
<pnp:Security
      BreakRoleInheritance="xsd:boolean"
      CopyRoleAssignments="xsd:boolean"
      ClearSubscopes="xsd:boolean">
   <pnp:AdditionalAdministrators />
   <pnp:AdditionalOwners />
   <pnp:AdditionalMembers />
   <pnp:AdditionalVisitors />
   <pnp:SiteGroups />
   <pnp:Permissions />
</pnp:Security>
```


Here follow the available child elements for the Security element.


Element|Type|Description
-------|----|-----------
AdditionalAdministrators|[UsersList](#userslist)|List of additional Administrators for the Site, optional collection of elements.
AdditionalOwners|[UsersList](#userslist)|List of additional Owners for the Site, optional collection of elements.
AdditionalMembers|[UsersList](#userslist)|List of additional Members for the Site, optional collection of elements.
AdditionalVisitors|[UsersList](#userslist)|List of additional Visitors for the Site, optional collection of elements.
SiteGroups|[SiteGroups](#sitegroups)|List of additional Groups for the Site, optional collection of elements.
Permissions|[Permissions](#permissions)|

Here follow the available attributes for the Security element.


Attibute|Type|Description
--------|----|-----------
BreakRoleInheritance|xsd:boolean|Declares whether the to break role inheritance for the site, if it is a sub-site, optional attribute.
CopyRoleAssignments|xsd:boolean|Defines whether to copy role assignments or not while breaking role inheritance, optional attribute.
ClearSubscopes|xsd:boolean|Defines whether to clear subscopes or not while breaking role inheritance for the site, optional attribute.
<a name="permissions"></a>
### Permissions


```xml
<pnp:Permissions>
   <pnp:RoleDefinitions />
   <pnp:RoleAssignments />
</pnp:Permissions>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
RoleDefinitions|[RoleDefinitions](#roledefinitions)|List of Role Definitions for the Site, optional collection of elements.
RoleAssignments|[RoleAssignments](#roleassignments)|List of Role Assignments for the Site, optional collection of elements.
<a name="navigation"></a>
### Navigation
The Navigation configurations of the Provisioning Template, optional collection of elements.

```xml
<pnp:Navigation
      EnableTreeView="xsd:boolean"
      AddNewPagesToNavigation="xsd:boolean"
      CreateFriendlyUrlsForNewPages="xsd:boolean">
   <pnp:GlobalNavigation />
   <pnp:CurrentNavigation />
</pnp:Navigation>
```


Here follow the available child elements for the Navigation element.


Element|Type|Description
-------|----|-----------
GlobalNavigation|[GlobalNavigation](#globalnavigation)|The Global Navigation settings for the Provisioning Template, optional element.
CurrentNavigation|[CurrentNavigation](#currentnavigation)|The Current Navigation settings for the Provisioning Template, optional element.

Here follow the available attributes for the Navigation element.


Attibute|Type|Description
--------|----|-----------
EnableTreeView|xsd:boolean|Declares whether the tree view has to be enabled at the site level or not, optional attribute.
AddNewPagesToNavigation|xsd:boolean|Declares whether the New Page ribbon command will automatically create a navigation item for the newly created page, optional attribute.
CreateFriendlyUrlsForNewPages|xsd:boolean|Declares whether the New Page ribbon command will automatically create a friendly URL for the newly created page, optional attribute.
<a name="globalnavigation"></a>
### GlobalNavigation
The Global Navigation settings for the Provisioning Template, optional element.

```xml
<pnp:GlobalNavigation
      NavigationType="">
   <pnp:StructuralNavigation />
   <pnp:ManagedNavigation />
</pnp:GlobalNavigation>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
StructuralNavigation|[StructuralNavigation](#structuralnavigation)|
ManagedNavigation|[ManagedNavigation](#managednavigation)|

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
NavigationType||Defines the type of Global Navigation, required attribute.
<a name="currentnavigation"></a>
### CurrentNavigation
The Current Navigation settings for the Provisioning Template, optional element.

```xml
<pnp:CurrentNavigation
      NavigationType="">
   <pnp:StructuralNavigation />
   <pnp:ManagedNavigation />
</pnp:CurrentNavigation>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
StructuralNavigation|[StructuralNavigation](#structuralnavigation)|
ManagedNavigation|[ManagedNavigation](#managednavigation)|

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
NavigationType||Defines the type of Current Navigation, required attribute.
<a name="managednavigation"></a>
### ManagedNavigation
Defines the Managed Navigation settings of a site, optional element.

```xml
<pnp:ManagedNavigation
      TermStoreId="pnp:ReplaceableString"
      TermSetId="pnp:ReplaceableString">
</pnp:ManagedNavigation>
```


Here follow the available attributes for the ManagedNavigation element.


Attibute|Type|Description
--------|----|-----------
TermStoreId|ReplaceableString|Defines the TermStore ID for the Managed Navigation, required attribute.
TermSetId|ReplaceableString|Defines the TermSet ID for the Managed Navigation, required attribute.
<a name="structuralnavigation"></a>
### StructuralNavigation
Defines the Structural Navigation settings of a site.

```xml
<pnp:StructuralNavigation
      RemoveExistingNodes="xsd:boolean">
   <pnp:NavigationNode />
</pnp:StructuralNavigation>
```


Here follow the available child elements for the StructuralNavigation element.


Element|Type|Description
-------|----|-----------
NavigationNode|[NavigationNode](#navigationnode)|

Here follow the available attributes for the StructuralNavigation element.


Attibute|Type|Description
--------|----|-----------
RemoveExistingNodes|xsd:boolean|Defines whether to remove existing nodes before creating those described through this element, required attribute.
<a name="navigationnode"></a>
### NavigationNode
Defines a Navigation Node for the Structural Navigation of a site.

```xml
<pnp:NavigationNode
      Title="pnp:ReplaceableString"
      Url="pnp:ReplaceableString"
      IsExternal="xsd:boolean"
      IsVisible="xsd:boolean">
   <pnp:NavigationNode />
</pnp:NavigationNode>
```


Here follow the available child elements for the NavigationNode element.


Element|Type|Description
-------|----|-----------
NavigationNode|[NavigationNode](#navigationnode)|

Here follow the available attributes for the NavigationNode element.


Attibute|Type|Description
--------|----|-----------
Title|ReplaceableString|Defines the Title of a Navigation Node for the Structural Navigation of a site.
Url|ReplaceableString|Defines the Url of a Navigation Node for the Structural Navigation of a site.
IsExternal|xsd:boolean|Defines whether the Navigation Node for the Structural Navigation targets an External resource.
IsVisible|xsd:boolean|Defines whether the Navigation Node for the Structural Navigation is visible or not.
<a name="features"></a>
### Features
The Features (Site or Web) to activate or deactivate while applying the Provisioning Template, optional collection of elements.

```xml
<pnp:Features>
   <pnp:SiteFeatures />
   <pnp:WebFeatures />
</pnp:Features>
```


Here follow the available child elements for the Features element.


Element|Type|Description
-------|----|-----------
SiteFeatures|[FeaturesList](#featureslist)|The Site Features to activate or deactivate while applying the Provisioning Template, optional collection of elements.
WebFeatures|[FeaturesList](#featureslist)|The Web Features to activate or deactivate while applying the Provisioning Template, optional collection of elements.
<a name="customactions"></a>
### CustomActions
The Custom Actions (Site or Web) to provision with the Provisioning Template, optional element.

```xml
<pnp:CustomActions>
   <pnp:SiteCustomActions />
   <pnp:WebCustomActions />
</pnp:CustomActions>
```


Here follow the available child elements for the CustomActions element.


Element|Type|Description
-------|----|-----------
SiteCustomActions|[CustomActionsList](#customactionslist)|The Site Custom Actions to provision while applying the Provisioning Template, optional collection of elements.
WebCustomActions|[CustomActionsList](#customactionslist)|The Web Custom Actions to provision while applying the Provisioning Template, optional collection of elements.
<a name="pages"></a>
### Pages
The Pages to provision into the target Site through the Provisioning Template, optional collection of elements.

```xml
<pnp:Pages>
   <pnp:Page />
</pnp:Pages>
```


Here follow the available child elements for the Pages element.


Element|Type|Description
-------|----|-----------
Page|[Page](#page)|
<a name="propertybagentry"></a>
### PropertyBagEntry
The Property Bag Entry of the Provisioning Template.

```xml
<pnp:PropertyBagEntry
      Overwrite="xsd:boolean"
      Indexed="xsd:boolean">
</pnp:PropertyBagEntry>
```


Here follow the available attributes for the PropertyBagEntry element.


Attibute|Type|Description
--------|----|-----------
Overwrite|xsd:boolean|Declares whether the Property Bag Entry has to overwrite an already existing entry, optional attribute.
Indexed|xsd:boolean|Declares whether the Property Bag Entry has to be indexed, optional attribute.
<a name="stringdictionaryitem"></a>
### StringDictionaryItem
Defines a StringDictionary element.

```xml
<pnp:StringDictionaryItem
      Key="xsd:string"
      Value="xsd:string">
</pnp:StringDictionaryItem>
```


Here follow the available attributes for the StringDictionaryItem element.


Attibute|Type|Description
--------|----|-----------
Key|xsd:string|The Key of the property to store in the StringDictionary, required attribute.
Value|xsd:string|The Value of the property to store in the StringDictionary, required attribute.
<a name="userslist"></a>
### UsersList
List of Users for the Site Security, collection of elements.

```xml
<pnp:UsersList>
   <pnp:User />
</pnp:UsersList>
```


Here follow the available child elements for the UsersList element.


Element|Type|Description
-------|----|-----------
User|[User](#user)|
<a name="user"></a>
### User
The base type for a User element.

```xml
<pnp:User
      Name="xsd:string">
</pnp:User>
```


Here follow the available attributes for the User element.


Attibute|Type|Description
--------|----|-----------
Name|xsd:string|The Name of the User, required attribute.
<a name="sitegroups"></a>
### SiteGroups
List of Site Groups for the Site Security, collection of elements.

```xml
<pnp:SiteGroups>
   <pnp:SiteGroup />
</pnp:SiteGroups>
```


Here follow the available child elements for the SiteGroups element.


Element|Type|Description
-------|----|-----------
SiteGroup|[SiteGroup](#sitegroup)|
<a name="sitegroup"></a>
### SiteGroup
The base type for a Site Group element.

```xml
<pnp:SiteGroup
      Title="xsd:string"
      Description="xsd:string"
      Owner="xsd:string"
      AllowMembersEditMembership="xsd:boolean"
      AllowRequestToJoinLeave="xsd:boolean"
      AutoAcceptRequestToJoinLeave="xsd:boolean"
      OnlyAllowMembersViewMembership="xsd:boolean"
      RequestToJoinLeaveEmailSetting="xsd:string">
   <pnp:Members />
</pnp:SiteGroup>
```


Here follow the available child elements for the SiteGroup element.


Element|Type|Description
-------|----|-----------
Members|[UsersList](#userslist)|The list of members of the Site Group, optional element.

Here follow the available attributes for the SiteGroup element.


Attibute|Type|Description
--------|----|-----------
Title|xsd:string|The Title of the Site Group, required attribute.
Description|xsd:string|The Description of the Site Group, optional attribute.
Owner|xsd:string|The Owner of the Site Group, required attribute.
AllowMembersEditMembership|xsd:boolean|Defines whether the members can edit membership of the Site Group, optional attribute.
AllowRequestToJoinLeave|xsd:boolean|Defines whether to allow requests to join or leave the Site Group, optional attribute.
AutoAcceptRequestToJoinLeave|xsd:boolean|Defines whether to auto-accept requests to join or leave the Site Group, optional attribute.
OnlyAllowMembersViewMembership|xsd:boolean|Defines whether to allow members only to view the membership of the Site Group, optional attribute.
RequestToJoinLeaveEmailSetting|xsd:string|Defines the email address used for membership requests to join or leave will be sent for the Site Group, optional attribute.
<a name="roledefinitions"></a>
### RoleDefinitions
List of Role Definitions for a target RoleAssignment, collection of elements.

```xml
<pnp:RoleDefinitions>
   <pnp:RoleDefinition />
</pnp:RoleDefinitions>
```


Here follow the available child elements for the RoleDefinitions element.


Element|Type|Description
-------|----|-----------
RoleDefinition|[RoleDefinition](#roledefinition)|
<a name="roledefinition"></a>
### RoleDefinition


```xml
<pnp:RoleDefinition
      Name="xsd:string"
      Description="xsd:string">
   <pnp:Permissions />
</pnp:RoleDefinition>
```


Here follow the available child elements for the RoleDefinition element.


Element|Type|Description
-------|----|-----------
Permissions|[Permissions](#permissions)|Defines the Permissions of the Role Definition, required element.

Here follow the available attributes for the RoleDefinition element.


Attibute|Type|Description
--------|----|-----------
Name|xsd:string|Defines the Name of the Role Definition, required attribute.
Description|xsd:string|Defines the Description of the Role Definition, optional attribute.
<a name="permissions"></a>
### Permissions
Defines the Permissions of the Role Definition, required element.

```xml
<pnp:Permissions>
   <pnp:Permission />
</pnp:Permissions>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Permission|[Permission](#permission)|Defines a Permission for a Role Definition.
<a name="roleassignments"></a>
### RoleAssignments
List of Role Assignments for a target Principal, collection of elements.

```xml
<pnp:RoleAssignments>
   <pnp:RoleAssignment />
</pnp:RoleAssignments>
```


Here follow the available child elements for the RoleAssignments element.


Element|Type|Description
-------|----|-----------
RoleAssignment|[RoleAssignment](#roleassignment)|
<a name="roleassignment"></a>
### RoleAssignment


```xml
<pnp:RoleAssignment
      Principal="xsd:string"
      RoleDefinition="xsd:string"
      Remove="xsd:boolean">
</pnp:RoleAssignment>
```


Here follow the available attributes for the RoleAssignment element.


Attibute|Type|Description
--------|----|-----------
Principal|xsd:string|Defines the Role to which the assignment will apply, required attribute.
RoleDefinition|xsd:string|Defines the Role to which the assignment will apply, required attribute.
Remove|xsd:boolean|Allows to remove a role assignment, instead of adding it. It is an optional attribute, and by default it assumes a value of false.
<a name="objectsecurity"></a>
### ObjectSecurity
Defines a set of Role Assignments for specific principals.

```xml
<pnp:ObjectSecurity>
   <pnp:BreakRoleInheritance />
</pnp:ObjectSecurity>
```


Here follow the available child elements for the ObjectSecurity element.


Element|Type|Description
-------|----|-----------
BreakRoleInheritance|[BreakRoleInheritance](#breakroleinheritance)|
<a name="breakroleinheritance"></a>
### BreakRoleInheritance
Declares a section of custom permissions, breaking role inheritance from parent.

```xml
<pnp:BreakRoleInheritance
      CopyRoleAssignments="xsd:boolean"
      ClearSubscopes="xsd:boolean">
   <pnp:RoleAssignment />
</pnp:BreakRoleInheritance>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
RoleAssignment|[RoleAssignment](#roleassignment)|

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
CopyRoleAssignments|xsd:boolean|Defines whether to copy role assignments or not while breaking role inheritance, required attribute.
ClearSubscopes|xsd:boolean|Defines whether to clear subscopes or not while breaking role inheritance, required attribute.
<a name="listinstance"></a>
### ListInstance
Defines a ListInstance element

```xml
<pnp:ListInstance
      Title="xsd:string"
      Description="xsd:string"
      DocumentTemplate="xsd:string"
      OnQuickLaunch="xsd:boolean"
      TemplateType="xsd:int"
      Url="xsd:string"
      ForceCheckout="xsd:boolean"
      EnableVersioning="xsd:boolean"
      EnableMinorVersions="xsd:boolean"
      EnableModeration="xsd:boolean"
      MinorVersionLimit="xsd:int"
      MaxVersionLimit="xsd:int"
      DraftVersionVisibility="xsd:int"
      RemoveExistingContentTypes="xsd:boolean"
      TemplateFeatureID="pnp:GUID"
      ContentTypesEnabled="xsd:boolean"
      Hidden="xsd:boolean"
      EnableAttachments="xsd:boolean"
      EnableFolderCreation="xsd:boolean"
      NoCrawl="xsd:boolean"
      ListExperience=""
      DefaultDisplayFormUrl="xsd:string"
      DefaultEditFormUrl="xsd:string"
      DefaultNewFormUrl="xsd:string"
      Direction=""
      ImageUrl="xsd:string"
      IrmExpire="xsd:boolean"
      IrmReject="xsd:boolean"
      IsApplicationList="xsd:boolean"
      ReadSecurity="xsd:int"
      WriteSecurity="xsd:int"
      ValidationFormula="xsd:string"
      ValidationMessage="xsd:string">
   <pnp:ContentTypeBindings />
   <pnp:Views />
   <pnp:Fields />
   <pnp:FieldRefs />
   <pnp:DataRows />
   <pnp:Folders />
   <pnp:FieldDefaults />
   <pnp:Security />
   <pnp:UserCustomActions />
   <pnp:Webhooks />
   <pnp:IRMSettings />
   <pnp:DataSource />
</pnp:ListInstance>
```


Here follow the available child elements for the ListInstance element.


Element|Type|Description
-------|----|-----------
ContentTypeBindings|[ContentTypeBindings](#contenttypebindings)|The ContentTypeBindings entries of the List Instance, optional collection of elements.
Views|[Views](#views)|The Views entries of the List Instance, optional collection of elements.
Fields|[Fields](#fields)|The Fields entries of the List Instance, optional collection of elements.
FieldRefs|[FieldRefs](#fieldrefs)|The FieldRefs entries of the List Instance, optional collection of elements.
DataRows|[DataRows](#datarows)|Defines a collection of rows that will be added to the List Instance, optional element.
Folders|[Folders](#folders)|Defines a collection of folders (eventually nested) that will be provisioned into the target list/library, optional element.
FieldDefaults|[FieldDefaults](#fielddefaults)|Defines a list of default values for the Fields of the List Instance, optional collection of elements.
Security|[ObjectSecurity](#objectsecurity)|Defines the Security rules for the List Instance, optional element.
UserCustomActions|[CustomActionsList](#customactionslist)|Defines any Custom Action for the List Instance, optional element.
Webhooks|[WebhooksList](#webhookslist)|Defines any Webhook for the current list instance.
IRMSettings|[IRMSettings](#irmsettings)|Declares the Information Rights Management settings for the list or library.
DataSource|[DataSource](#datasource)|Allows defining the Data Source for an external list, optional element.

Here follow the available attributes for the ListInstance element.


Attibute|Type|Description
--------|----|-----------
Title|xsd:string|The Title of the List Instance, required attribute.
Description|xsd:string|The Description of the List Instance, optional attribute.
DocumentTemplate|xsd:string|The DocumentTemplate of the List Instance, optional attribute.
OnQuickLaunch|xsd:boolean|The OnQuickLaunch flag for the List Instance, optional attribute.
TemplateType|xsd:int|The TemplateType of the List Instance, required attribute. Values available here: https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.listtemplatetype.aspx
Url|xsd:string|The Url of the List Instance, required attribute.
ForceCheckout|xsd:boolean|The ForceCheckout flag for the List Instance, optional attribute.
EnableVersioning|xsd:boolean|The EnableVersioning flag for the List Instance, optional attribute.
EnableMinorVersions|xsd:boolean|The EnableMinorVersions flag for the List Instance, optional attribute.
EnableModeration|xsd:boolean|The EnableModeration flag for the List Instance, optional attribute.
MinorVersionLimit|xsd:int|The MinorVersionLimit for versions history for the List Instance, optional attribute.
MaxVersionLimit|xsd:int|The MaxVersionLimit for versions history for the List Instance, optional attribute.
DraftVersionVisibility|xsd:int|The DraftVersionVisibility for the List Instance, optional attribute. The property will be cast to enum DraftVersionVisibility 0 - Reader - Any user who can read items, 1 - Author - Only users who can edit items, 2 - Approver - Only users who can approve items (and the author of the item)
RemoveExistingContentTypes|xsd:boolean|The RemoveExistingContentTypes flag for the List Instance, optional attribute.
TemplateFeatureID|GUID|The TemplateFeatureID for the feature on which the List Instance is based, optional attribute.
ContentTypesEnabled|xsd:boolean|The ContentTypesEnabled flag for the List Instance, optional attribute.
Hidden|xsd:boolean|The Hidden flag for the List Instance, optional attribute.
EnableAttachments|xsd:boolean|The EnableAttachments flag for the List Instance, optional attribute.
EnableFolderCreation|xsd:boolean|The EnableFolderCreation flag for the List Instance, optional attribute.
NoCrawl|xsd:boolean|Defines if the current list or library has to be included in crawling, optional attribute.
ListExperience||Defines the current list UI/UX experience (valid for SPO only).
DefaultDisplayFormUrl|xsd:string|Defines a value that specifies the location of the default display form for the list.
DefaultEditFormUrl|xsd:string|Defines a value that specifies the URL of the edit form to use for list items in the list.
DefaultNewFormUrl|xsd:string|Defines a value that specifies the location of the default new form for the list.
Direction||Defines a value that specifies the reading order of the list.
ImageUrl|xsd:string|Defines a value that specifies the URI for the icon of the list, optional attribute.
IrmExpire|xsd:boolean|Defines if IRM Expire property, optional attribute.
IrmReject|xsd:boolean|Defines the IRM Reject property, optional attribute.
IsApplicationList|xsd:boolean|Defines a value that specifies a flag that a client application can use to determine whether to display the list, optional attribute.
ReadSecurity|xsd:int|Defines the Read Security property, optional attribute.
WriteSecurity|xsd:int|Defines the Write Security property, optional attribute.
ValidationFormula|xsd:string|Defines a value that specifies the data validation criteria for a list item, optional attribute.
ValidationMessage|xsd:string|Defines a value that specifies the error message returned when data validation fails for a list item, optional attribute.
<a name="contenttypebindings"></a>
### ContentTypeBindings
The ContentTypeBindings entries of the List Instance, optional collection of elements.

```xml
<pnp:ContentTypeBindings>
   <pnp:ContentTypeBinding />
</pnp:ContentTypeBindings>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
ContentTypeBinding|[ContentTypeBinding](#contenttypebinding)|
<a name="views"></a>
### Views
The Views entries of the List Instance, optional collection of elements.

```xml
<pnp:Views
      RemoveExistingViews="xsd:boolean">
   <!-- Any other XML content -->
</pnp:Views>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
RemoveExistingViews|xsd:boolean|A flag to declare if the existing views of the List Instance have to be removed, before adding the custom views, optional attribute.
<a name="fields"></a>
### Fields
The Fields entries of the List Instance, optional collection of elements.

```xml
<pnp:Fields>
   <!-- Any other XML content -->
</pnp:Fields>
```

<a name="fieldrefs"></a>
### FieldRefs
The FieldRefs entries of the List Instance, optional collection of elements.

```xml
<pnp:FieldRefs>
   <pnp:FieldRef />
</pnp:FieldRefs>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
FieldRef|[ListInstanceFieldRef](#listinstancefieldref)|
<a name="datarows"></a>
### DataRows
Defines a collection of rows that will be added to the List Instance, optional element.

```xml
<pnp:DataRows
      KeyColumn="xsd:string"
      UpdateBehavior="">
   <pnp:DataRow />
</pnp:DataRows>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
DataRow|[DataRow](#datarow)|

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
KeyColumn|xsd:string|Optional attribute to declare the name of the Key Column, if any, used to identify any already existing DataRows. If this attribute has a value and the target list already has items with matching KeyColumn values, the engine will handle the matching DataRows based on the value of the UpdateBehavior attribute. If UpdateBehavior has a value of Skip, the DataRows will be skipped. If UpdateBehavior has a value of Overwrite, the existing items will be updated with the values defined in the DataRows.
UpdateBehavior||If the DataRow already exists on target list, this attribute defines whether the DataRow will be overwritten or skipped.
<a name="folders"></a>
### Folders
Defines a collection of folders (eventually nested) that will be provisioned into the target list/library, optional element.

```xml
<pnp:Folders>
   <pnp:Folder />
</pnp:Folders>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Folder|[Folder](#folder)|
<a name="fielddefaults"></a>
### FieldDefaults
Defines a list of default values for the Fields of the List Instance, optional collection of elements.

```xml
<pnp:FieldDefaults>
   <pnp:FieldDefault />
</pnp:FieldDefaults>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
FieldDefault|[FieldDefault](#fielddefault)|Defines a default value for a Field of the List Instance.
<a name="datasource"></a>
### DataSource
Allows defining the Data Source for an external list, optional element.

```xml
<pnp:DataSource>
   <pnp:DataSourceItem />
</pnp:DataSource>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
DataSourceItem|[StringDictionaryItem](#stringdictionaryitem)|A single data source property for an external list.
<a name="irmsettings"></a>
### IRMSettings
Declares the Information Rights Management settings for the list or library.

```xml
<pnp:IRMSettings
      Enabled="xsd:boolean"
      AllowPrint="xsd:boolean"
      AllowScript="xsd:boolean"
      AllowWriteCopy="xsd:boolean"
      DisableDocumentBrowserView="xsd:boolean"
      DocumentAccessExpireDays="xsd:int"
      DocumentLibraryProtectionExpiresInDays="pnp:ReplaceableInt"
      EnableDocumentAccessExpire="xsd:boolean"
      EnableDocumentBrowserPublishingView="xsd:boolean"
      EnableGroupProtection="xsd:boolean"
      EnableLicenseCacheExpire="xsd:boolean"
      GroupName="xsd:string"
      LicenseCacheExpireDays="xsd:int"
      PolicyDescription="xsd:string"
      PolicyTitle="xsd:string">
</pnp:IRMSettings>
```


Here follow the available attributes for the IRMSettings element.


Attibute|Type|Description
--------|----|-----------
Enabled|xsd:boolean|Defines whether the IRM settings have to be enabled or not.
AllowPrint|xsd:boolean|Defines whether a viewer can print the downloaded document.
AllowScript|xsd:boolean|Defines whether a viewer can run a script on the downloaded document.
AllowWriteCopy|xsd:boolean|Defines whether a viewer can write on a copy of the downloaded document.
DisableDocumentBrowserView|xsd:boolean|Defines whether to block Office Web Application Companion applications (WACs) from showing this document.
DocumentAccessExpireDays|xsd:int|Defines the number of days after which the downloaded document will expire.
DocumentLibraryProtectionExpiresInDays|ReplaceableInt|Defines the expire days for the Information Rights Management (IRM) protection of this document library will expire.
EnableDocumentAccessExpire|xsd:boolean|Defines whether the downloaded document will expire.
EnableDocumentBrowserPublishingView|xsd:boolean|Defines whether to enable Office Web Application Companion applications (WACs) to publishing view.
EnableGroupProtection|xsd:boolean|Defines whether the permission of the downloaded document is applicable to a group.
EnableLicenseCacheExpire|xsd:boolean|Defines whether a user must verify their credentials after some interval.
GroupName|xsd:string|Defines the group name (email address) that the permission is also applicable to.
LicenseCacheExpireDays|xsd:int|Defines the number of days that the application that opens the document caches the IRM license. When these elapse, the application will connect to the IRM server to validate the license.
PolicyDescription|xsd:string|Defines the permission policy description.
PolicyTitle|xsd:string|Defines the permission policy title.
<a name="folder"></a>
### Folder
Defines a folder that will be provisioned into the target list/library.

```xml
<pnp:Folder
      Name="xsd:string">
   <pnp:Folder />
   <pnp:Security />
   <pnp:PropertyBagEntries />
</pnp:Folder>
```


Here follow the available child elements for the Folder element.


Element|Type|Description
-------|----|-----------
Folder|[Folder](#folder)|A child Folder of another Folder item, optional element.
Security|[ObjectSecurity](#objectsecurity)|Defines the security rules for the row that will be added to the List Instance, optional element.
PropertyBagEntries|[PropertyBagEntries](#propertybagentries)|The Property Bag entries of the Folder, optional collection of elements.

Here follow the available attributes for the Folder element.


Attibute|Type|Description
--------|----|-----------
Name|xsd:string|The Name of the Folder, required attribute.
<a name="datavalue"></a>
### DataValue
The DataValue of a single field of a row to insert into a target ListInstance.

```xml
<pnp:DataValue>
</pnp:DataValue>
```

<a name="fielddefault"></a>
### FieldDefault
The FieldDefault of a single field of list or library for target ListInstance.

```xml
<pnp:FieldDefault>
</pnp:FieldDefault>
```

<a name="contenttype"></a>
### ContentType
Defines a Content Type.

```xml
<pnp:ContentType
      ID="pnp:ContentTypeId"
      Name="xsd:string"
      Description="xsd:string"
      Group="xsd:string"
      Hidden="xsd:boolean"
      Sealed="xsd:boolean"
      ReadOnly="xsd:boolean"
      Overwrite="xsd:boolean"
      NewFormUrl="xsd:string"
      EditFormUrl="xsd:string"
      DisplayFormUrl="xsd:string">
   <pnp:FieldRefs />
   <pnp:DocumentTemplate />
   <pnp:DocumentSetTemplate />
</pnp:ContentType>
```


Here follow the available child elements for the ContentType element.


Element|Type|Description
-------|----|-----------
FieldRefs|[FieldRefs](#fieldrefs)|The FieldRefs entries of the List Instance, optional collection of elements.
DocumentTemplate|[DocumentTemplate](#documenttemplate)|Specifies the document template for the content type. This is the file which SharePoint Foundation opens as a template when a user requests a new item of this content type.
DocumentSetTemplate|[DocumentSetTemplate](#documentsettemplate)|Specifies the properties of the DocumentSet Template if the ContentType defines a DocumentSet.

Here follow the available attributes for the ContentType element.


Attibute|Type|Description
--------|----|-----------
ID|ContentTypeId|The value of the Content Type ID, required attribute.
Name|xsd:string|The Name of the Content Type, required attribute.
Description|xsd:string|The Description of the Content Type, optional attribute.
Group|xsd:string|The Group of the Content Type, optional attribute.
Hidden|xsd:boolean|Optional Boolean. True to define the content type as hidden. If you define a content type as hidden, SharePoint Foundation does not display that content type on the New button in list views.
Sealed|xsd:boolean|Optional Boolean. True to prevent changes to this content type. You cannot change the value of this attribute through the user interface, but you can change it in code if you have sufficient rights. You must have site collection administrator rights to unseal a content type.
ReadOnly|xsd:boolean|Optional Boolean. TRUE to specify that the content type cannot be edited without explicitly removing the read-only setting. This can be done either in the user interface or in code.
Overwrite|xsd:boolean|Optional Boolean. TRUE to overwrite an existing content type with the same ID.
NewFormUrl|xsd:string|Specifies the URL of a custom new form to use for list items that have been assigned the content type, optional attribute.
EditFormUrl|xsd:string|Specifies the URL of a custom edit form to use for list items that have been assigned the content type, optional attribute.
DisplayFormUrl|xsd:string|Specifies the URL of a custom display form to use for list items that have been assigned the content type, optional attribute.
<a name="fieldrefs"></a>
### FieldRefs
The FieldRefs entries of the List Instance, optional collection of elements.

```xml
<pnp:FieldRefs>
   <pnp:FieldRef />
</pnp:FieldRefs>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
FieldRef|[ContentTypeFieldRef](#contenttypefieldref)|
<a name="documenttemplate"></a>
### DocumentTemplate
Specifies the document template for the content type. This is the file which SharePoint Foundation opens as a template when a user requests a new item of this content type.

```xml
<pnp:DocumentTemplate
      TargetName="xsd:string">
</pnp:DocumentTemplate>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
TargetName|xsd:string|The value of the Content Type ID, required attribute.
<a name="contenttypebinding"></a>
### ContentTypeBinding
Defines the binding between a ListInstance and a ContentType.

```xml
<pnp:ContentTypeBinding
      ContentTypeID="pnp:ContentTypeId"
      Default="xsd:boolean"
      Remove="xsd:boolean"
      Hidden="xsd:boolean">
</pnp:ContentTypeBinding>
```


Here follow the available attributes for the ContentTypeBinding element.


Attibute|Type|Description
--------|----|-----------
ContentTypeID|ContentTypeId|The value of the Content Type ID to bind, required attribute.
Default|xsd:boolean|Declares if the Content Type should be the default Content Type in the list or library, optional attribute.
Remove|xsd:boolean|Declares if the Content Type should be Removed from the list or library, optional attribute.
Hidden|xsd:boolean|Declares if the Content Type should be Hidden from New button of the list or library, optional attribute.
<a name="documentsettemplate"></a>
### DocumentSetTemplate
Defines a DocumentSet Template for creating multiple DocumentSet instances.

```xml
<pnp:DocumentSetTemplate
      WelcomePage="xsd:string">
   <pnp:AllowedContentTypes />
   <pnp:DefaultDocuments />
   <pnp:SharedFields />
   <pnp:WelcomePageFields />
   <pnp:XmlDocuments />
</pnp:DocumentSetTemplate>
```


Here follow the available child elements for the DocumentSetTemplate element.


Element|Type|Description
-------|----|-----------
AllowedContentTypes|[AllowedContentTypes](#allowedcontenttypes)|
DefaultDocuments|[DefaultDocuments](#defaultdocuments)|
SharedFields|[SharedFields](#sharedfields)|
WelcomePageFields|[WelcomePageFields](#welcomepagefields)|
XmlDocuments|[XmlDocuments](#xmldocuments)|Defines any custom XmlDocument section for the DocumentSet, it is optional.

Here follow the available attributes for the DocumentSetTemplate element.


Attibute|Type|Description
--------|----|-----------
WelcomePage|xsd:string|Defines the custom WelcomePage for the Document Set, optional attribute.
<a name="allowedcontenttypes"></a>
### AllowedContentTypes
The list of allowed Content Types for the Document Set, optional element.

```xml
<pnp:AllowedContentTypes
      RemoveExistingContentTypes="xsd:boolean">
   <pnp:AllowedContentType />
</pnp:AllowedContentTypes>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
AllowedContentType|[AllowedContentType](#allowedcontenttype)|

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
RemoveExistingContentTypes|xsd:boolean|The RemoveExistingContentTypes flag for the Allowed Content Types of the current Document Set, optional attribute.
<a name="defaultdocuments"></a>
### DefaultDocuments
The list of default Documents for the Document Set, optional element.

```xml
<pnp:DefaultDocuments>
   <pnp:DefaultDocument />
</pnp:DefaultDocuments>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
DefaultDocument|[DefaultDocument](#defaultdocument)|
<a name="sharedfields"></a>
### SharedFields
The list of Shared Fields for the Document Set, optional element.

```xml
<pnp:SharedFields>
   <pnp:SharedField />
</pnp:SharedFields>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
SharedField|[DocumentSetFieldRef](#documentsetfieldref)|
<a name="welcomepagefields"></a>
### WelcomePageFields
The list of Welcome Page Fields for the Document Set, optional element.

```xml
<pnp:WelcomePageFields>
   <pnp:WelcomePageField />
</pnp:WelcomePageFields>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
WelcomePageField|[DocumentSetFieldRef](#documentsetfieldref)|
<a name="xmldocuments"></a>
### XmlDocuments
Defines any custom XmlDocument section for the DocumentSet, it is optional.

```xml
<pnp:XmlDocuments>
   <!-- Any other XML content -->
</pnp:XmlDocuments>
```

<a name="featureslist"></a>
### FeaturesList
Defines a collection of elements of type Feature.

```xml
<pnp:FeaturesList>
   <pnp:Feature />
</pnp:FeaturesList>
```


Here follow the available child elements for the FeaturesList element.


Element|Type|Description
-------|----|-----------
Feature|[Feature](#feature)|
<a name="feature"></a>
### Feature
Defines a single Site or Web Feature, which will be activated or deactivated while applying the Provisioning Template.

```xml
<pnp:Feature
      ID="pnp:GUID"
      Deactivate="xsd:boolean"
      Description="xsd:string">
</pnp:Feature>
```


Here follow the available attributes for the Feature element.


Attibute|Type|Description
--------|----|-----------
ID|GUID|The unique ID of the Feature, required attribute.
Deactivate|xsd:boolean|Defines if the feature has to be deactivated or activated while applying the Provisioning Template, optional attribute.
Description|xsd:string|The Description of the feature, optional attribute.
<a name="fieldrefbase"></a>
### FieldRefBase


```xml
<pnp:FieldRefBase
      ID="pnp:ReplaceableString">
</pnp:FieldRefBase>
```


Here follow the available attributes for the FieldRefBase element.


Attibute|Type|Description
--------|----|-----------
ID|ReplaceableString|The value of the field ID to bind, required attribute.
<a name="fieldreffull"></a>
### FieldRefFull


```xml
<pnp:FieldRefFull>
</pnp:FieldRefFull>
```

<a name="listinstancefieldref"></a>
### ListInstanceFieldRef
Defines the binding between a ListInstance and a Field.

```xml
<pnp:ListInstanceFieldRef
      DisplayName="xsd:string"
      Remove="xsd:boolean">
</pnp:ListInstanceFieldRef>
```


Here follow the available attributes for the ListInstanceFieldRef element.


Attibute|Type|Description
--------|----|-----------
DisplayName|xsd:string|The display name of the field to bind, only applicable to fields that will be added to lists, optional attribute.
Remove|xsd:boolean|Declares if the FieldRef should be Removed from the list or library, optional attribute.
<a name="contenttypefieldref"></a>
### ContentTypeFieldRef
Defines the binding between a ContentType and a Field.

```xml
<pnp:ContentTypeFieldRef>
</pnp:ContentTypeFieldRef>
```

<a name="documentsetfieldref"></a>
### DocumentSetFieldRef
Defines the binding between a Document Set and a Field.

```xml
<pnp:DocumentSetFieldRef>
</pnp:DocumentSetFieldRef>
```

<a name="customactionslist"></a>
### CustomActionsList
Defines a collection of elements of type CustomAction.

```xml
<pnp:CustomActionsList>
   <pnp:CustomAction />
</pnp:CustomActionsList>
```


Here follow the available child elements for the CustomActionsList element.


Element|Type|Description
-------|----|-----------
CustomAction|[CustomAction](#customaction)|
<a name="customaction"></a>
### CustomAction
Defines a Custom Action, which will be provisioned while applying the Provisioning Template.

```xml
<pnp:CustomAction
      Name="xsd:string"
      Description="xsd:string"
      Group="xsd:string"
      Location="xsd:string"
      Title="xsd:string"
      Sequence="xsd:int"
      Rights="xsd:string"
      Url="xsd:string"
      Enabled="xsd:boolean"
      Remove="xsd:boolean"
      ScriptBlock="xsd:string"
      ImageUrl="xsd:string"
      ScriptSrc="xsd:string"
      RegistrationId="xsd:string"
      RegistrationType="pnp:RegistrationType"
      ClientSideComponentId="xsd:string"
      ClientSideComponentProperties="xsd:string">
   <pnp:CommandUIExtension />
</pnp:CustomAction>
```


Here follow the available child elements for the CustomAction element.


Element|Type|Description
-------|----|-----------
CommandUIExtension|[CommandUIExtension](#commanduiextension)|Defines the Custom UI Extension XML, optional element.

Here follow the available attributes for the CustomAction element.


Attibute|Type|Description
--------|----|-----------
Name|xsd:string|The Name of the CustomAction, required attribute.
Description|xsd:string|The Description of the CustomAction, optional attribute.
Group|xsd:string|The Group of the CustomAction, optional attribute.
Location|xsd:string|The Location of the CustomAction, required attribute.
Title|xsd:string|The Title of the CustomAction, required attribute.
Sequence|xsd:int|The Sequence of the CustomAction, optional attribute.
Rights|xsd:string|The Rights for the CustomAction, based on values from Microsoft.SharePoint.Client.PermissionKind, optional attribute.
Url|xsd:string|The URL of the CustomAction, optional attribute.
Enabled|xsd:boolean|The Enabled flag for the CustomAction, optional attribute.
Remove|xsd:boolean|To Remove a CustomAction, optional attribute.
ScriptBlock|xsd:string|The ScriptBlock of the CustomAction, optional attribute.
ImageUrl|xsd:string|The ImageUrl of the CustomAction, optional attribute.
ScriptSrc|xsd:string|The ScriptSrc of the CustomAction, optional attribute.
RegistrationId|xsd:string|The RegistrationId of the CustomAction, optional attribute.
RegistrationType|RegistrationType|The RegistrationType of the CustomAction, optional attribute.
ClientSideComponentId|xsd:string|The Client-Side Component Id of a customizer, optional attribute.
ClientSideComponentProperties|xsd:string|The Client-Side Component Properties of a customizer, optional attribute.
<a name="commanduiextension"></a>
### CommandUIExtension
Defines the Custom UI Extension XML, optional element.

```xml
<pnp:CommandUIExtension>
   <!-- Any other XML content -->
</pnp:CommandUIExtension>
```

<a name="sitewebhookslist"></a>
### SiteWebhooksList
Defines a collection of elements of type SiteWebhook.

```xml
<pnp:SiteWebhooksList>
   <pnp:Webhook />
</pnp:SiteWebhooksList>
```


Here follow the available child elements for the SiteWebhooksList element.


Element|Type|Description
-------|----|-----------
Webhook|[SiteWebhook](#sitewebhook)|
<a name="webhookslist"></a>
### WebhooksList
Defines a collection of elements of type Webhook.

```xml
<pnp:WebhooksList>
   <pnp:Webhook />
</pnp:WebhooksList>
```


Here follow the available child elements for the WebhooksList element.


Element|Type|Description
-------|----|-----------
Webhook|[Webhook](#webhook)|
<a name="webhook"></a>
### Webhook
Defines a single element of type Webhook.

```xml
<pnp:Webhook
      ServerNotificationUrl="pnp:ReplaceableString"
      ExpiresInDays="pnp:ReplaceableInt">
</pnp:Webhook>
```


Here follow the available attributes for the Webhook element.


Attibute|Type|Description
--------|----|-----------
ServerNotificationUrl|ReplaceableString|The Server Notification URL of the Webhook, required attribute.
ExpiresInDays|ReplaceableInt|The expire days for the subscription of the Webhook, required attribute.
<a name="sitewebhook"></a>
### SiteWebhook
Defines a single element of type SiteWebhook.

```xml
<pnp:SiteWebhook
      SiteWebhookType="">
</pnp:SiteWebhook>
```


Here follow the available attributes for the SiteWebhook element.


Attibute|Type|Description
--------|----|-----------
SiteWebhookType||
<a name="clientsidepages"></a>
### ClientSidePages
Defines a collection of elements of type ClientSidePage.

```xml
<pnp:ClientSidePages>
   <pnp:ClientSidePage />
</pnp:ClientSidePages>
```


Here follow the available child elements for the ClientSidePages element.


Element|Type|Description
-------|----|-----------
ClientSidePage|[ClientSidePage](#clientsidepage)|
<a name="clientsidepage"></a>
### ClientSidePage
Defines a single element of type ClientSidePage.

```xml
<pnp:ClientSidePage
      PageName="xsd:string"
      PromoteAsNewsArticle="xsd:boolean"
      Overwrite="xsd:boolean"
      Layout="xsd:string"
      Publish="xsd:boolean"
      EnableComments="xsd:boolean"
      Title="xsd:string">
   <pnp:Sections />
   <pnp:Header />
</pnp:ClientSidePage>
```


Here follow the available child elements for the ClientSidePage element.


Element|Type|Description
-------|----|-----------
Sections|[Sections](#sections)|Defines the Canvas sections for a single ClientSidePage.
Header|[Header](#header)|Defines the layout of the Header for the current client side page

Here follow the available attributes for the ClientSidePage element.


Attibute|Type|Description
--------|----|-----------
PageName|xsd:string|Defines the page name for a single ClientSidePage.
PromoteAsNewsArticle|xsd:boolean|Declares to promote the page as a news article.
Overwrite|xsd:boolean|Can the page be overwritten if it exists.
Layout|xsd:string|Defines the target layout for the client-side page, optional attribute (default: Article).
Publish|xsd:boolean|Defines whether the page will be published or not, optional attribute (default: true).
EnableComments|xsd:boolean|Defines whether the page will have comments enabled or not, optional attribute (default: true).
Title|xsd:string|Defines the Title of the page, optional attribute.
<a name="sections"></a>
### Sections
Defines the Canvas sections for a single ClientSidePage.

```xml
<pnp:Sections>
   <pnp:Section />
</pnp:Sections>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Section|[CanvasSection](#canvassection)|Defines a Canvas section for a single ClientSidePage.
<a name="header"></a>
### Header
Defines the layout of the Header for the current client side page

```xml
<pnp:Header
      Type=""
      ServerRelativeImageUrl="xsd:string"
      TranslateX="xsd:double"
      TranslateY="xsd:double">
</pnp:Header>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Type||Defines the layout of the Header for the current client side page
ServerRelativeImageUrl|xsd:string|Defines the server-relative URL of the image for the header of the current client side page.
TranslateX|xsd:double|Defines the x-translate of the image for the header of the current client side page.
TranslateY|xsd:double|Defines the y-translate of of the image for the header of the current client side page.
<a name="canvassection"></a>
### CanvasSection
A Canvas Section for a Client-side Page.

```xml
<pnp:CanvasSection
      Order="xsd:int"
      Type="">
   <pnp:Controls />
</pnp:CanvasSection>
```


Here follow the available child elements for the CanvasSection element.


Element|Type|Description
-------|----|-----------
Controls|[Controls](#controls)|A collection of Canvas Controls for a Client-side Page.

Here follow the available attributes for the CanvasSection element.


Attibute|Type|Description
--------|----|-----------
Order|xsd:int|The order of the Canvas Section for a Client-side Page.
Type||The type of the Canvas Section for a Client-side Page.
<a name="controls"></a>
### Controls
A collection of Canvas Controls for a Client-side Page.

```xml
<pnp:Controls>
   <pnp:CanvasControl />
</pnp:Controls>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
CanvasControl|[CanvasControl](#canvascontrol)|A single Canvas Control for a Client-side Page.
<a name="canvascontrol"></a>
### CanvasControl
Defines a Canvas Control for a Client-side Page.

```xml
<pnp:CanvasControl
      WebPartType=""
      CustomWebPartName="xsd:string"
      JsonControlData="xsd:string"
      ControlId="pnp:GUID"
      Order="xsd:int"
      Column="xsd:int">
   <pnp:CanvasControlProperties />
</pnp:CanvasControl>
```


Here follow the available child elements for the CanvasControl element.


Element|Type|Description
-------|----|-----------
CanvasControlProperties|[CanvasControlProperties](#canvascontrolproperties)|Custom properties for the client-side web part control.

Here follow the available attributes for the CanvasControl element.


Attibute|Type|Description
--------|----|-----------
WebPartType||The Type of Client-side Web Part.
CustomWebPartName|xsd:string|The Name of the client-side web part if the WebPartType attribute has a value of "Custom".
JsonControlData|xsd:string|The JSON Control Data for Canvas Control of a Client-side Page.
ControlId|GUID|The Instance Id for Canvas Control of a Client-side Page.
Order|xsd:int|The order of the Canvas Control for a Client-side Page.
Column|xsd:int|The Column of the Section in which the Canvas Control will be inserted. Optional, default 0.
<a name="canvascontrolproperties"></a>
### CanvasControlProperties
Custom properties for the client-side web part control.

```xml
<pnp:CanvasControlProperties>
   <pnp:CanvasControlProperty />
</pnp:CanvasControlProperties>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
CanvasControlProperty|[StringDictionaryItem](#stringdictionaryitem)|A single property of a client-side web part control.
<a name="canvascolumn"></a>
### CanvasColumn
Defines a Canvas Section for a Client-side Page.

```xml
<pnp:CanvasColumn
      Order="xsd:int"
      ColumnFactor="xsd:int">
   <pnp:Controls />
</pnp:CanvasColumn>
```


Here follow the available child elements for the CanvasColumn element.


Element|Type|Description
-------|----|-----------
Controls|[Controls](#controls)|A collection of Canvas Controls for a Client-side Page.

Here follow the available attributes for the CanvasColumn element.


Attibute|Type|Description
--------|----|-----------
Order|xsd:int|The order of the Canvas section for a Client-side Page.
ColumnFactor|xsd:int|The column Factor for Canvas column of a Client-side Page.
<a name="controls"></a>
### Controls
A collection of Canvas Controls for a Client-side Page.

```xml
<pnp:Controls>
   <pnp:CanvasControl />
</pnp:Controls>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
CanvasControl|[CanvasControl](#canvascontrol)|A single Canvas Control for a Client-side Page.
<a name="fileproperties"></a>
### FileProperties
A collection of File Properties.

```xml
<pnp:FileProperties>
   <pnp:Property />
</pnp:FileProperties>
```


Here follow the available child elements for the FileProperties element.


Element|Type|Description
-------|----|-----------
Property|[StringDictionaryItem](#stringdictionaryitem)|
<a name="file"></a>
### File
Defines a File element, to describe a file that will be provisioned into the target Site.

```xml
<pnp:File
      Src="xsd:string"
      Folder="xsd:string"
      Overwrite="xsd:boolean"
      Level="pnp:FileLevel"
      TargetFileName="pnp:ReplaceableString">
   <pnp:Properties />
   <pnp:WebParts />
   <pnp:Security />
</pnp:File>
```


Here follow the available child elements for the File element.


Element|Type|Description
-------|----|-----------
Properties|[FileProperties](#fileproperties)|The File Properties, optional collection of elements.
WebParts|[WebParts](#webparts)|The webparts to add to the page, optional collection of elements.
Security|[ObjectSecurity](#objectsecurity)|

Here follow the available attributes for the File element.


Attibute|Type|Description
--------|----|-----------
Src|xsd:string|The Src of the File, required attribute.
Folder|xsd:string|The TargetFolder of the File, required attribute.
Overwrite|xsd:boolean|The Overwrite flag for the File, optional attribute.
Level|FileLevel|The Level status for the File, optional attribute.
TargetFileName|ReplaceableString|The Target file name for the File, optional attribute. If missing, the original file name will be used.
<a name="webparts"></a>
### WebParts
The webparts to add to the page, optional collection of elements.

```xml
<pnp:WebParts>
   <pnp:WebPart />
</pnp:WebParts>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
WebPart|[WebPartPageWebPart](#webpartpagewebpart)|
<a name="directory"></a>
### Directory
Defines a Directory element, to describe a folder in the current repository that will be used to upload files into the target Site.

```xml
<pnp:Directory
      Src="pnp:ReplaceableString"
      Folder="pnp:ReplaceableString"
      Overwrite="xsd:boolean"
      Level="pnp:FileLevel"
      Recursive="xsd:boolean"
      IncludedExtensions="xsd:string"
      ExcludedExtensions="xsd:string"
      MetadataMappingFile="pnp:ReplaceableString">
   <pnp:Security />
</pnp:Directory>
```


Here follow the available child elements for the Directory element.


Element|Type|Description
-------|----|-----------
Security|[ObjectSecurity](#objectsecurity)|

Here follow the available attributes for the Directory element.


Attibute|Type|Description
--------|----|-----------
Src|ReplaceableString|The Src of the Directory, required attribute.
Folder|ReplaceableString|The TargetFolder of the Directory, required attribute.
Overwrite|xsd:boolean|The Overwrite flag for the File items in the Directory, optional attribute.
Level|FileLevel|The Level status for the File, optional attribute.
Recursive|xsd:boolean|Defines whether to recursively browse through all the child folders of the Directory, optional attribute.
IncludedExtensions|xsd:string|The file Extensions (lower case) to include while uploading the Directory, optional attribute.
ExcludedExtensions|xsd:string|The file Extensions (lower case) to exclude while uploading the Directory, optional attribute.
MetadataMappingFile|ReplaceableString|The file path of JSON mapping file with metadata for files to upload in the Directory, optional attribute.
<a name="page"></a>
### Page
Defines a Page element, to describe a page that will be provisioned into the target Site. Because of the Layout attribute, the assumption is made that you're referring/creating a WikiPage.

```xml
<pnp:Page
      Url="xsd:string"
      Overwrite="xsd:boolean"
      Layout="pnp:WikiPageLayout">
   <pnp:WebParts />
   <pnp:Fields />
   <pnp:Security />
</pnp:Page>
```


Here follow the available child elements for the Page element.


Element|Type|Description
-------|----|-----------
WebParts|[WebParts](#webparts)|The webparts to add to the page, optional collection of elements.
Fields|[Fields](#fields)|The Fields to setup for the Page, optional collection of elements.
Security|[ObjectSecurity](#objectsecurity)|

Here follow the available attributes for the Page element.


Attibute|Type|Description
--------|----|-----------
Url|xsd:string|The server relative url of the page, supports tokens, required attribute.
Overwrite|xsd:boolean|If set, overwrites an existing page in the case of a wikipage, optional attribute.
Layout|WikiPageLayout|Defines the layout of the wikipage, required attribute.
<a name="webparts"></a>
### WebParts
The webparts to add to the page, optional collection of elements.

```xml
<pnp:WebParts>
   <pnp:WebPart />
</pnp:WebParts>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
WebPart|[WikiPageWebPart](#wikipagewebpart)|
<a name="fields"></a>
### Fields
The Fields to setup for the Page, optional collection of elements.

```xml
<pnp:Fields>
   <pnp:Field />
</pnp:Fields>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Field|[BaseFieldValue](#basefieldvalue)|
<a name="wikipagewebpart"></a>
### WikiPageWebPart
Defines a WebPart to be added to a WikiPage.

```xml
<pnp:WikiPageWebPart
      Title="xsd:string"
      Row="xsd:int"
      Column="xsd:int">
   <pnp:Contents />
</pnp:WikiPageWebPart>
```


Here follow the available child elements for the WikiPageWebPart element.


Element|Type|Description
-------|----|-----------
Contents|[Contents](#contents)|Defines the WebPart XML, required element.

Here follow the available attributes for the WikiPageWebPart element.


Attibute|Type|Description
--------|----|-----------
Title|xsd:string|Defines the title of the WebPart, required attribute.
Row|xsd:int|Defines the row to add the WebPart to, required attribute.
Column|xsd:int|Defines the column to add the WebPart to, required attribute.
<a name="contents"></a>
### Contents
Defines the WebPart XML, required element.

```xml
<pnp:Contents>
   <!-- Any other XML content -->
</pnp:Contents>
```

<a name="webpartpagewebpart"></a>
### WebPartPageWebPart
Defines a webpart to be added to a WebPart Page.

```xml
<pnp:WebPartPageWebPart
      Title="xsd:string"
      Zone="xsd:string"
      Order="xsd:int">
   <pnp:Contents />
</pnp:WebPartPageWebPart>
```


Here follow the available child elements for the WebPartPageWebPart element.


Element|Type|Description
-------|----|-----------
Contents|[Contents](#contents)|Defines the WebPart XML, required element.

Here follow the available attributes for the WebPartPageWebPart element.


Attibute|Type|Description
--------|----|-----------
Title|xsd:string|Defines the title of the WebPart, required attribute.
Zone|xsd:string|Defines the zone of a WebPart Page to add the webpart to, required attribute.
Order|xsd:int|Defines the index of the WebPart in the zone, required attribute.
<a name="contents"></a>
### Contents
Defines the WebPart XML, required element.

```xml
<pnp:Contents>
   <!-- Any other XML content -->
</pnp:Contents>
```

<a name="composedlook"></a>
### ComposedLook
Defines a ComposedLook element.

```xml
<pnp:ComposedLook
      Name="xsd:string"
      ColorFile="xsd:string"
      FontFile="xsd:string"
      BackgroundFile="xsd:string"
      Version="xsd:int">
</pnp:ComposedLook>
```


Here follow the available attributes for the ComposedLook element.


Attibute|Type|Description
--------|----|-----------
Name|xsd:string|The Name of the ComposedLook, required attribute.
ColorFile|xsd:string|The ColorFile of the ComposedLook, required attribute.
FontFile|xsd:string|The FontFile of the ComposedLook, required attribute.
BackgroundFile|xsd:string|The BackgroundFile of the ComposedLook, optional attribute.
Version|xsd:int|The Version of the ComposedLook, optional attribute.
<a name="workflows"></a>
### Workflows
Defines the Workflows to provision.

```xml
<pnp:Workflows>
   <pnp:WorkflowDefinitions />
   <pnp:WorkflowSubscriptions />
</pnp:Workflows>
```


Here follow the available child elements for the Workflows element.


Element|Type|Description
-------|----|-----------
WorkflowDefinitions|[WorkflowDefinitions](#workflowdefinitions)|
WorkflowSubscriptions|[WorkflowSubscriptions](#workflowsubscriptions)|
<a name="workflowdefinitions"></a>
### WorkflowDefinitions
Defines the Workflows Definitions to provision.

```xml
<pnp:WorkflowDefinitions>
   <pnp:WorkflowDefinition />
</pnp:WorkflowDefinitions>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
WorkflowDefinition|[WorkflowDefinition](#workflowdefinition)|
<a name="workflowsubscriptions"></a>
### WorkflowSubscriptions
Defines the Workflows Subscriptions to provision.

```xml
<pnp:WorkflowSubscriptions>
   <pnp:WorkflowSubscription />
</pnp:WorkflowSubscriptions>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
WorkflowSubscription|[WorkflowSubscription](#workflowsubscription)|
<a name="applicationlifecyclemanagement"></a>
### ApplicationLifecycleManagement
Defines the SharePoint Add-ins and SharePoint Framework solutions to provision, collection of elements.

```xml
<pnp:ApplicationLifecycleManagement>
   <pnp:AppCatalog />
   <pnp:Apps />
</pnp:ApplicationLifecycleManagement>
```


Here follow the available child elements for the ApplicationLifecycleManagement element.


Element|Type|Description
-------|----|-----------
AppCatalog|[AppCatalog](#appcatalog)|The App Catalog local to the current Site Collection
Apps|[Apps](#apps)|The SharePoint Add-ins and SharePoint Framework solutions to provision into the target Site through the ALM API of SharePoint Online, optional element.
<a name="apps"></a>
### Apps
The SharePoint Add-ins and SharePoint Framework solutions to provision into the target Site through the ALM API of SharePoint Online, optional element.

```xml
<pnp:Apps>
   <pnp:App />
</pnp:Apps>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
App|[App](#app)|
<a name="appcatalog"></a>
### AppCatalog
Element to manage the AppCatalog tenant-wide, or local to a specific Site Collection

```xml
<pnp:AppCatalog>
   <pnp:Package />
</pnp:AppCatalog>
```


Here follow the available child elements for the AppCatalog element.


Element|Type|Description
-------|----|-----------
Package|[Package](#package)|
<a name="package"></a>
### Package


```xml
<pnp:Package
      PackageId="pnp:ReplaceableString"
      Src="xsd:string"
      Action=""
      SkipFeatureDeployment="xsd:boolean"
      Overwrite="xsd:boolean">
</pnp:Package>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
PackageId|ReplaceableString|Defines the ID of the package to manage, optional attribute (either this one or the Src has to be provided)
Src|xsd:string|Defines the path of the package to manage, optional attribute (either this one or the PackageId has to be provided).
Action||Defines the action to execute against the package, required attribute
SkipFeatureDeployment|xsd:boolean|Defines whether to skip the feature deployment for tenant-wide enabled packages, optional attribute (default: false).
Overwrite|xsd:boolean|Defines whether to overwrite an already existing package in the AppCatalog, optional attribute (default: false).
<a name="contentdeliverynetwork"></a>
### ContentDeliveryNetwork
Element to manage the tenant-wide CDN.

```xml
<pnp:ContentDeliveryNetwork>
   <pnp:Public />
   <pnp:Private />
</pnp:ContentDeliveryNetwork>
```


Here follow the available child elements for the ContentDeliveryNetwork element.


Element|Type|Description
-------|----|-----------
Public|[CdnSetting](#cdnsetting)|Defines the Public CDN settings.
Private|[CdnSetting](#cdnsetting)|Defines the Private CDN settings.
<a name="cdnsetting"></a>
### CdnSetting
Defines the settings for a Public or Private CDN.

```xml
<pnp:CdnSetting
      Enabled="xsd:boolean">
   <pnp:Origins />
   <pnp:IncludeFileExtensions />
   <pnp:ExcludeRestrictedSiteClassifications />
   <pnp:ExcludeIfNoScriptDisabled />
</pnp:CdnSetting>
```


Here follow the available child elements for the CdnSetting element.


Element|Type|Description
-------|----|-----------
Origins|[Origins](#origins)|Defines the CDN Origins for the current CDN.
IncludeFileExtensions|xsd:string|Defines the file extensions to include in the CDN policy.
ExcludeRestrictedSiteClassifications|xsd:string|Defines the site classifications to exclude of the wild card origins.
ExcludeIfNoScriptDisabled|xsd:string|Allows to opt-out of sites that have disabled NoScript.

Here follow the available attributes for the CdnSetting element.


Attibute|Type|Description
--------|----|-----------
Enabled|xsd:boolean|Defines whether the CDN has to be enabled or disabled, required attribute.
<a name="origins"></a>
### Origins
Defines the CDN Origins for the current CDN.

```xml
<pnp:Origins>
   <pnp:Origin />
</pnp:Origins>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Origin|[Origin](#origin)|Defines a single CDN Origin for the current CDN.
<a name="publishing"></a>
### Publishing
Defines the Publishing configuration to provision.

```xml
<pnp:Publishing
      AutoCheckRequirements="">
   <pnp:DesignPackage />
   <pnp:AvailableWebTemplates />
   <pnp:PageLayouts />
   <pnp:ImageRenditions />
</pnp:Publishing>
```


Here follow the available child elements for the Publishing element.


Element|Type|Description
-------|----|-----------
DesignPackage|[DesignPackage](#designpackage)|Defines a Design Package to import into the current Publishing site, optional element.
AvailableWebTemplates|[AvailableWebTemplates](#availablewebtemplates)|Defines the Available Web Templates for the current Publishing site, optional collection of elements.
PageLayouts|[PageLayouts](#pagelayouts)|Defines the Available Page Layouts for the current Publishing site, optional collection of elements.
ImageRenditions|[ImageRenditions](#imagerenditions)|Defines the Image Renditions for the current Publishing site, optional collection of elements.

Here follow the available attributes for the Publishing element.


Attibute|Type|Description
--------|----|-----------
AutoCheckRequirements||Defines how an engine should behave if the requirements for provisioning publishing capabilities are not satisfied by the target site, required attribute.
<a name="designpackage"></a>
### DesignPackage
Defines a Design Package to import into the current Publishing site, optional element.

```xml
<pnp:DesignPackage
      DesignPackagePath="xsd:string"
      MajorVersion="xsd:int"
      MinorVersion="xsd:int"
      PackageGuid="pnp:GUID"
      PackageName="xsd:string">
</pnp:DesignPackage>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
DesignPackagePath|xsd:string|Defines the path of the Design Package to import into the current Publishing site, required attribute.
MajorVersion|xsd:int|The Major Version of the Design Package to import into the current Publishing site, optional attribute.
MinorVersion|xsd:int|The Minor Version of the Design Package to import into the current Publishing site, optional attribute.
PackageGuid|GUID|The ID of the Design Package to import into the current Publishing site, optional attribute.
PackageName|xsd:string|The Name of the Design Package to import into the current Publishing site, required attribute.
<a name="availablewebtemplates"></a>
### AvailableWebTemplates
Defines the Available Web Templates for the current Publishing site, optional collection of elements.

```xml
<pnp:AvailableWebTemplates>
   <pnp:WebTemplate />
</pnp:AvailableWebTemplates>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
WebTemplate|[WebTemplate](#webtemplate)|Defines an available Web Template for the current Publishing site.
<a name="pagelayouts"></a>
### PageLayouts
Defines the Available Page Layouts for the current Publishing site, optional collection of elements.

```xml
<pnp:PageLayouts
      Default="xsd:string">
   <pnp:PageLayout />
</pnp:PageLayouts>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
PageLayout|[PageLayout](#pagelayout)|Defines an available Page Layout for the current Publishing site.

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Default|xsd:string|Defines the URL of the Default Page Layout for the current Publishing site, if any. Optional attribute.
<a name="imagerenditions"></a>
### ImageRenditions
Defines the Image Renditions for the current Publishing site, optional collection of elements.

```xml
<pnp:ImageRenditions>
   <pnp:ImageRendition />
</pnp:ImageRenditions>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
ImageRendition|[ImageRendition](#imagerendition)|Defines an available Image Rendition for the current Publishing site.
<a name="provider"></a>
### Provider
Defines an Extensibility Provider.

```xml
<pnp:Provider
      Enabled="xsd:boolean"
      HandlerType="xsd:string">
   <pnp:Configuration />
</pnp:Provider>
```


Here follow the available child elements for the Provider element.


Element|Type|Description
-------|----|-----------
Configuration|[Configuration](#configuration)|Defines an optional configuration section for the Extensibility Provider. The configuration section can be any XML.

Here follow the available attributes for the Provider element.


Attibute|Type|Description
--------|----|-----------
Enabled|xsd:boolean|Defines whether the Extensibility Provider is enabled or not, optional attribute.
HandlerType|xsd:string|The type of the handler. It can be a FQN of a .NET type, the URL of a node.js file, or whatever else, required attribute.
<a name="configuration"></a>
### Configuration
Defines an optional configuration section for the Extensibility Provider. The configuration section can be any XML.

```xml
<pnp:Configuration>
   <!-- Any other XML content -->
</pnp:Configuration>
```

<a name="provisioningtemplatefile"></a>
### ProvisioningTemplateFile
An element that references an external file.

```xml
<pnp:ProvisioningTemplateFile
      File="xsd:string"
      ID="xsd:ID">
</pnp:ProvisioningTemplateFile>
```


Here follow the available attributes for the ProvisioningTemplateFile element.


Attibute|Type|Description
--------|----|-----------
File|xsd:string|Absolute or relative path to the file, required attribute.
ID|xsd:ID|ID of the referenced template, required attribute.
<a name="provisioningtemplatereference"></a>
### ProvisioningTemplateReference
An element that references an external file.

```xml
<pnp:ProvisioningTemplateReference
      ID="xsd:IDREF">
</pnp:ProvisioningTemplateReference>
```


Here follow the available attributes for the ProvisioningTemplateReference element.


Attibute|Type|Description
--------|----|-----------
ID|xsd:IDREF|ID of the referenced template, required attribute.
<a name="sequence"></a>
### Sequence
Each Provisioning file is split into a set of Sequence elements. The Sequence element groups the artefacts to be provisioned into groups. The Sequences must be evaluated by the provisioning engine in the order in which they appear.

```xml
<pnp:Sequence
      SequenceType=""
      ID="xsd:ID">
   <pnp:SiteCollection />
   <pnp:Site />
   <pnp:TermStore />
   <pnp:Extensions />
</pnp:Sequence>
```


Here follow the available child elements for the Sequence element.


Element|Type|Description
-------|----|-----------
SiteCollection|[SiteCollection](#sitecollection)|A Site Collection to provision through a Sequence, optional element.
Site|[Site](#site)|A Site to provision through a Sequence, optional element.
TermStore|[TermStore](#termstore)|A Term Store to provision through a Sequence, optional element.
Extensions|[Extensions](#extensions)|Any Extension to provision through a Sequence, optional element.

Here follow the available attributes for the Sequence element.


Attibute|Type|Description
--------|----|-----------
SequenceType||Instructions to the Provisioning engine on how the Containers within the Sequence can be provisioned.
ID|xsd:ID|A unique identifier of the Sequence, required attribute.
<a name="sitecollection"></a>
### SiteCollection
Defines a SiteCollection that will be created into the target tenant/farm.

```xml
<pnp:SiteCollection
      Url="pnp:ReplaceableString">
   <pnp:Templates />
</pnp:SiteCollection>
```


Here follow the available child elements for the SiteCollection element.


Element|Type|Description
-------|----|-----------
Templates|[Templates](#templates)|Templates that can be provisioned together with the Site Collection, optional collection of elements.

Here follow the available attributes for the SiteCollection element.


Attibute|Type|Description
--------|----|-----------
Url|ReplaceableString|Absolute Url to the site, required attribute.
<a name="site"></a>
### Site
Defines a Site that will be created into a target Site Collection.

```xml
<pnp:Site
      UseSamePermissionsAsParentSite="xsd:boolean"
      Url="pnp:ReplaceableString">
   <pnp:Templates />
</pnp:Site>
```


Here follow the available child elements for the Site element.


Element|Type|Description
-------|----|-----------
Templates|[Templates](#templates)|Templates that can be provisioned together with the Site, optional collection of elements.

Here follow the available attributes for the Site element.


Attibute|Type|Description
--------|----|-----------
UseSamePermissionsAsParentSite|xsd:boolean|
Url|ReplaceableString|Relative Url to the site, required attribute.
<a name="termstore"></a>
### TermStore
A TermStore to use for provisioning of TermGroups.

```xml
<pnp:TermStore
      Scope="">
   <pnp:TermGroup />
</pnp:TermStore>
```


Here follow the available child elements for the TermStore element.


Element|Type|Description
-------|----|-----------
TermGroup|[TermGroup](#termgroup)|The TermGroup element to provision into the target TermStore through, optional collection of elements.

Here follow the available attributes for the TermStore element.


Attibute|Type|Description
--------|----|-----------
Scope||The scope of the term store, required attribute.
<a name="termgroup"></a>
### TermGroup
A TermGroup to use for provisioning of TermSets and Terms.

```xml
<pnp:TermGroup
      Description="xsd:string"
      SiteCollectionTermGroup="xsd:boolean"
      UpdateBehavior=""
      Name="xsd:string"
      ID="pnp:GUID">
</pnp:TermGroup>
```


Here follow the available attributes for the TermGroup element.


Attibute|Type|Description
--------|----|-----------
Description|xsd:string|The Description of the TermGroup to use for provisioning of TermSets and Terms, optional attribute.
SiteCollectionTermGroup|xsd:boolean|Declares if the TermGroup is the Site Collection Term Group, optional attribute.
UpdateBehavior||If the TermGroup already exists on target, this attribute defines whether the TermGroup will be overwritten or skipped.
Name|xsd:string|The Name of the Taxonomy Item, required attribute.
ID|GUID|The ID of the Taxonomy Item, optional attribute.
<a name="termsetitem"></a>
### TermSetItem
Base type for TermSets and Terms

```xml
<pnp:TermSetItem
      Owner="xsd:string"
      Description="xsd:string"
      IsAvailableForTagging="xsd:boolean">
</pnp:TermSetItem>
```


Here follow the available attributes for the TermSetItem element.


Attibute|Type|Description
--------|----|-----------
Owner|xsd:string|The Owner of the Term Set Item, optional attribute.
Description|xsd:string|The Description of the Term Set Item, optional attribute.
IsAvailableForTagging|xsd:boolean|Declares whether the Term Set Item is available for tagging, optional attribute.
<a name="termset"></a>
### TermSet
A TermSet to provision.

```xml
<pnp:TermSet
      Language="xsd:int"
      IsOpenForTermCreation="xsd:boolean">
</pnp:TermSet>
```


Here follow the available attributes for the TermSet element.


Attibute|Type|Description
--------|----|-----------
Language|xsd:int|The reference Language for the Term Set, optional attribute.
IsOpenForTermCreation|xsd:boolean|Declares whether the Term Set is open for terms creation or not, optional attribute.
<a name="term"></a>
### Term
A Term to provision into a TermSet or a hyerarchical Term.

```xml
<pnp:Term
      Language="xsd:int"
      CustomSortOrder="xsd:int"
      IsReused="xsd:boolean"
      IsSourceTerm="xsd:boolean"
      IsDeprecated="xsd:boolean"
      SourceTermId="pnp:GUID">
</pnp:Term>
```


Here follow the available attributes for the Term element.


Attibute|Type|Description
--------|----|-----------
Language|xsd:int|The reference Language for the Term, optional attribute.
CustomSortOrder|xsd:int|The Custom Sort Order for the Term, optional attribute.
IsReused|xsd:boolean|Declares if this term is reused, optional attribute.
IsSourceTerm|xsd:boolean|If the IsReused property is set to false, the current Term is not reused and this property will always be true. If the current Term is reused (IsReused returns true), then this property should be set to true if it is the source Term.
IsDeprecated|xsd:boolean|Declares if this term is deprecated, optional attribute.
SourceTermId|GUID|The ID of the source term if this term is reused, optional attribute.
<a name="taxonomyitemproperties"></a>
### TaxonomyItemProperties
A collection of Term Properties.

```xml
<pnp:TaxonomyItemProperties>
   <pnp:Property />
</pnp:TaxonomyItemProperties>
```


Here follow the available child elements for the TaxonomyItemProperties element.


Element|Type|Description
-------|----|-----------
Property|[StringDictionaryItem](#stringdictionaryitem)|A Term Property, collection of elements.
<a name="termlabels"></a>
### TermLabels
A collection of Term Labels, in order to support multi-language terms.

```xml
<pnp:TermLabels>
   <pnp:Label />
</pnp:TermLabels>
```


Here follow the available child elements for the TermLabels element.


Element|Type|Description
-------|----|-----------
Label|[Label](#label)|
<a name="label"></a>
### Label


```xml
<pnp:Label
      Language="xsd:int"
      Value="xsd:string"
      IsDefaultForLanguage="xsd:boolean">
</pnp:Label>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Language|xsd:int|The reference Language for the Term Label, required attribute.
Value|xsd:string|The Value for the Term Label, required attribute.
IsDefaultForLanguage|xsd:boolean|Declares whether the current Label is the default for the specific Language, optional attribute.
<a name="termsets"></a>
### TermSets
A collection of TermSets to provision.

```xml
<pnp:TermSets>
   <pnp:TermSet />
</pnp:TermSets>
```


Here follow the available child elements for the TermSets element.


Element|Type|Description
-------|----|-----------
TermSet|[TermSet](#termset)|A Term Set, optional collection of elements.
<a name="extensions"></a>
### Extensions
Extensions are custom XML elements and instructions that can be extensions of this default schema or vendor or engine specific extensions.

```xml
<pnp:Extensions>
   <!-- Any other XML content -->
</pnp:Extensions>
```

<a name="importsequence"></a>
### ImportSequence
Imports sequences from an external file. All current properties should be sent to that file.

```xml
<pnp:ImportSequence
      File="xsd:string">
</pnp:ImportSequence>
```


Here follow the available attributes for the ImportSequence element.


Attibute|Type|Description
--------|----|-----------
File|xsd:string|Absolute or relative path to the file, required attribute.
