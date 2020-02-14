
# PnP Provisioning Schema
----------
* Topic automatically generated on 11/25/2019*

## Namespace
The namespace of the PnP Provisioning Schema is:

http://schemas.dev.office.com/PnP/2019/09/ProvisioningSchema

All the elements have to be declared with that namespace reference.

## Root Elements
Here follows the list of root elements available in the PnP Provisioning Schema.
  
<a name="provisioning"></a>
### Provisioning


```xml
<pnp:Provisioning
      xmlns:pnp="http://schemas.dev.office.com/PnP/2019/09/ProvisioningSchema"
      Version="xsd:decimal"
      Author="xsd:string"
      Generator="xsd:string"
      ImagePreviewUrl="xsd:string"
      DisplayName="xsd:string"
      Description="xsd:string">
   <pnp:Preferences />
   <pnp:Localizations />
   <pnp:Tenant />
   <pnp:Templates />
   <pnp:Sequence />
   <pnp:Teams />
   <pnp:AzureActiveDirectory />
   <pnp:Drive />
   <pnp:ProvisioningWebhooks />
</pnp:Provisioning>
```


Here follow the available child elements for the Provisioning element.


Element|Type|Description
-------|----|-----------
Preferences|[Preferences](#preferences)|Section of preferences for the current provisioning definition.
Localizations|[Localizations](#localizations)|An optional list of localizations files to include.
Tenant|[Tenant](#tenant)|Entry point to manage tenant-wide settings.
Templates|[Templates](#templates)|An optional section made of provisioning templates.
Sequence|[Sequence](#sequence)|An optional section made of provisioning sequences, which can include Sites, Site Collections, Taxonomies, Provisioning Templates, etc.
Teams|[Teams](#teams)|Entry point to manage Microsoft Teams provisioning.
AzureActiveDirectory|[AzureActiveDirectory](#azureactivedirectory)|Entry point to manage Microsoft Azure Active Directory provisioning.
Drive|[Drive](#drive)|Entry point to manage OneDrive for Business provisioning.
ProvisioningWebhooks|[ProvisioningWebhooks](#provisioningwebhooks)|Entry point to manage global provisioning webhooks.

Here follow the available attributes for the Provisioning element.


Attibute|Type|Description
--------|----|-----------
Version|xsd:decimal|The Version of the Provisioning Template, optional attribute.
Author|xsd:string|The Author of the Provisioning Template, optional attribute.
Generator|xsd:string|Name of the tool generating this Provisioning File, optional attribute.
ImagePreviewUrl|xsd:string|The Image Preview Url of the Provisioning File, optional attribute.
DisplayName|xsd:string|The Display Name of the Provisioning File, optional attribute.
Description|xsd:string|The Description of the Provisioning file, optional attribute.
<a name="provisioningtemplate"></a>
### ProvisioningTemplate
Represents the root element of the SharePoint Provisioning Template.

```xml
<pnp:ProvisioningTemplate
      xmlns:pnp="http://schemas.dev.office.com/PnP/2019/09/ProvisioningSchema"
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
   <pnp:SiteSettings />
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
   <pnp:Header />
   <pnp:Footer />
   <pnp:ProvisioningTemplateWebhooks />
   <pnp:Theme />
</pnp:ProvisioningTemplate>
```


Here follow the available child elements for the ProvisioningTemplate element.


Element|Type|Description
-------|----|-----------
Properties|[ProvisioningTemplateProperties](#provisioningtemplateproperties)|A set of custom Properties for the Provisioning Template, optional element.
SitePolicy|[ReplaceableString](#replaceablestring)|The Site Policy of the Provisioning Template, optional element.
WebSettings|[WebSettings](#websettings)|Section of Settings for the current Web Site, optional element.
SiteSettings|[SiteSettings](#sitesettings)|Section of Settings for the current Site Collection, optional element.
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
Header|[Header](#header)|The Header of the Site provisioned through the Provisioning Template, optional element.
Footer|[Footer](#footer)|The Footer of the Site provisioned through the Provisioning Template, optional element.
ProvisioningTemplateWebhooks|[ProvisioningTemplateWebhooks](#provisioningtemplatewebhooks)|The Webhooks for the Provisioning Template, optional element.
Theme|[Theme](#theme)|The Theme for the Provisioning Template, optional element.

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
   <pnp:Themes />
   <pnp:SPUsersProfiles />
   <pnp:Office365GroupLifecyclePolicies />
   <pnp:Office365GroupsSettings />
</pnp:Tenant>
```


Here follow the available child elements for the Tenant element.


Element|Type|Description
-------|----|-----------
AppCatalog|[AppCatalog](#appcatalog)|Entry point for the tenant-wide AppCatalog, optional element.
WebApiPermissions|[WebApiPermissions](#webapipermissions)|Entry point for the tenant-wide Web API permissions, optional element.
ContentDeliveryNetwork|[ContentDeliveryNetwork](#contentdeliverynetwork)|Entry point for the tenant-wide Content Delivery Network, optional element.
SiteDesigns|[SiteDesigns](#sitedesigns)|Entry point for the tenant-wide Site Designs, optional element.
SiteScripts|[SiteScripts](#sitescripts)|Entry point for the tenant-wide Site Scripts, optional element.
StorageEntities|[StorageEntities](#storageentities)|Entry point for the tenant-wide properties (Storage Entities), optional element.
Themes|[Themes](#themes)|Entry point for the tenant-wide Themes, optional element.
SPUsersProfiles|[SPUsersProfiles](#spusersprofiles)|Entry point for the User Profile properties, optional element.
Office365GroupLifecyclePolicies|[Office365GroupLifecyclePolicies](#office365grouplifecyclepolicies)|Defines a collection of Group Lifecycle Policies, optional element.
Office365GroupsSettings|[Office365GroupsSettings](#office365groupssettings)|Defines the configuration properties for Unified Groups at tenant level, optional element.
<a name="office365groupssettings"></a>
### Office365GroupsSettings
Defines the configuration properties for Unified Groups at tenant level, optional element.

```xml
<pnp:Office365GroupsSettings>
   <pnp:Property />
</pnp:Office365GroupsSettings>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Property|[StringDictionaryItem](#stringdictionaryitem)|Defines a custom property and value for the Unified Group Settings
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
<a name="themes"></a>
### Themes
Collection of tenant-wide Themes

```xml
<pnp:Themes>
   <pnp:Theme />
</pnp:Themes>
```


Here follow the available child elements for the Themes element.


Element|Type|Description
-------|----|-----------
Theme|[Theme](#theme)|Defines a single tenant-wide Theme
<a name="theme"></a>
### Theme
Defines a single Theme

```xml
<pnp:Theme
      Name="xsd:string"
      IsInverted="xsd:boolean"
      Overwrite="xsd:boolean">
</pnp:Theme>
```


Here follow the available attributes for the Theme element.


Attibute|Type|Description
--------|----|-----------
Name|xsd:string|Defines the name of the tenant-wide Theme
IsInverted|xsd:boolean|Defines the name of the tenant-wide Theme
Overwrite|xsd:boolean|Defines whether to overwrite an already existing theme, optional attribute.
<a name="spusersprofiles"></a>
### SPUsersProfiles
Collection of UserProfile objects with custom properties

```xml
<pnp:SPUsersProfiles>
   <pnp:SPUserProfile />
</pnp:SPUsersProfiles>
```


Here follow the available child elements for the SPUsersProfiles element.


Element|Type|Description
-------|----|-----------
SPUserProfile|[SPUserProfile](#spuserprofile)|Defines a single UserProfile
<a name="spuserprofile"></a>
### SPUserProfile
Defines a single UserProfile

```xml
<pnp:SPUserProfile
      TargetUser="pnp:ReplaceableString"
      TargetGroup="pnp:ReplaceableString">
   <pnp:Property />
</pnp:SPUserProfile>
```


Here follow the available child elements for the SPUserProfile element.


Element|Type|Description
-------|----|-----------
Property|[StringDictionaryItem](#stringdictionaryitem)|Defines a custom property and value for a UserProfile

Here follow the available attributes for the SPUserProfile element.


Attibute|Type|Description
--------|----|-----------
TargetUser|ReplaceableString|Defines the target User
TargetGroup|ReplaceableString|Defines the target AAD Group
<a name="office365grouplifecyclepolicies"></a>
### Office365GroupLifecyclePolicies
Collection of Group Lifecycle Policies objects

```xml
<pnp:Office365GroupLifecyclePolicies>
   <pnp:Office365GroupLifecyclePolicy />
</pnp:Office365GroupLifecyclePolicies>
```


Here follow the available child elements for the Office365GroupLifecyclePolicies element.


Element|Type|Description
-------|----|-----------
Office365GroupLifecyclePolicy|[Office365GroupLifecyclePolicy](#office365grouplifecyclepolicy)|Defines a single Group Lifecycle Policy
<a name="office365grouplifecyclepolicy"></a>
### Office365GroupLifecyclePolicy
Defines a single Group Lifecycle Policy

```xml
<pnp:Office365GroupLifecyclePolicy
      ID="xsd:string"
      GroupLifetimeInDays="xsd:int"
      AlternateNotificationEmails="pnp:ReplaceableString"
      ManagedGroupTypes="">
</pnp:Office365GroupLifecyclePolicy>
```


Here follow the available attributes for the Office365GroupLifecyclePolicy element.


Attibute|Type|Description
--------|----|-----------
ID|xsd:string|Defines the PnP internal ID for the Group Lifetime Policy, required attribute.
GroupLifetimeInDays|xsd:int|Defines the Group Lifetime in Days for the policy, required attribute.
AlternateNotificationEmails|ReplaceableString|Defines the Alternate Notification Emails for the policy, required attribute. Separate multiple email addresses with a semicolon.
ManagedGroupTypes||Defines the Managed Group Types for the policy, required attribute.
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
<a name="provisioningtemplatewebhooks"></a>
### ProvisioningTemplateWebhooks
Allows to define one or more webhooks that will be invoked by the engine, upon completion of specific actions

```xml
<pnp:ProvisioningTemplateWebhooks>
   <pnp:ProvisioningTemplateWebhook />
</pnp:ProvisioningTemplateWebhooks>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
ProvisioningTemplateWebhook|[ProvisioningWebhook](#provisioningwebhook)|Defines a single Provisioning Template Webhook
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
      CommentsOnSitePagesDisabled="xsd:boolean"
      QuickLaunchEnabled="xsd:boolean"
      ExcludeFromOfflineClient="xsd:boolean"
      MembersCanShare="xsd:boolean"
      DisableFlows="xsd:boolean"
      DisableAppViews="xsd:boolean"
      HorizontalQuickLaunch="xsd:boolean"
      SearchScope=""
      IsMultilingual="xsd:boolean"
      OverwriteTranslationsOnChange="xsd:boolean">
   <pnp:AlternateUICultures />
</pnp:WebSettings>
```


Here follow the available child elements for the WebSettings element.


Element|Type|Description
-------|----|-----------
AlternateUICultures|[AlternateUICultures](#alternateuicultures)|Defines the list of Alternate UI Cultures for the current web, optional element.

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
CommentsOnSitePagesDisabled|xsd:boolean|Enables or disables comments on client side pages, optional attribute.
QuickLaunchEnabled|xsd:boolean|Enables or disables the QuickLaunch for the site, optional attribute.
ExcludeFromOfflineClient|xsd:boolean|Defines whether to exclude the web from offline client, optional attribute.
MembersCanShare|xsd:boolean|Defines whether members can share content from the current web, optional attribute.
DisableFlows|xsd:boolean|Defines whether disable flows for the current web, optional attribute.
DisableAppViews|xsd:boolean|Defines whether disable PowerApps for the current web, optional attribute.
HorizontalQuickLaunch|xsd:boolean|Defines whether to enable the Horizontal QuickLaunch for the current web, optional attribute.
SearchScope||Defines the SearchScope for the site, optional attribute.
IsMultilingual|xsd:boolean|Defines whether to enable Multilingual capabilities for the current web, optional attribute.
OverwriteTranslationsOnChange|xsd:boolean|Defines whether to OverwriteTranslationsOnChange on change for the current web, optional attribute.
<a name="alternateuicultures"></a>
### AlternateUICultures
Defines the list of Alternate UI Cultures for the current web, optional element.

```xml
<pnp:AlternateUICultures>
   <pnp:AlternateUICulture />
</pnp:AlternateUICultures>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
AlternateUICulture|[AlternateUICulture](#alternateuiculture)|Defines an Alternate UI Culture item for the current web, optional element.
<a name="sitesettings"></a>
### SiteSettings
Section of Settings for the current Site Collection, optional element.

```xml
<pnp:SiteSettings
      AllowDesigner="xsd:boolean"
      AllowCreateDeclarativeWorkflow="xsd:boolean"
      AllowSaveDeclarativeWorkflowAsTemplate="xsd:boolean"
      AllowSavePublishDeclarativeWorkflow="xsd:boolean"
      SocialBarOnSitePagesDisabled="xsd:boolean">
</pnp:SiteSettings>
```


Here follow the available attributes for the SiteSettings element.


Attibute|Type|Description
--------|----|-----------
AllowDesigner|xsd:boolean|Defines whether a designer can be used on this site collection, optional attribute.
AllowCreateDeclarativeWorkflow|xsd:boolean|Defines whether creation of declarative workflows is allowed in the site collection, optional attribute.
AllowSaveDeclarativeWorkflowAsTemplate|xsd:boolean|Defines whether saving of declarative workflows is allowed in the site collection, optional attribute.
AllowSavePublishDeclarativeWorkflow|xsd:boolean|Defines whether publishing of declarative workflows is allowed in the site collection, optional attribute.
SocialBarOnSitePagesDisabled|xsd:boolean|Defines whether social bar is disabled on Site Pages in this site collection, optional attribute.
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
      ResetRoleInheritance="xsd:boolean"
      CopyRoleAssignments="xsd:boolean"
      RemoveExistingUniqueRoleAssignments="xsd:boolean"
      ClearSubscopes="xsd:boolean"
      AssociatedGroups="pnp:ReplaceableString"
      AssociatedOwnerGroup="pnp:ReplaceableString"
      AssociatedMemberGroup="pnp:ReplaceableString"
      AssociatedVisitorGroup="pnp:ReplaceableString">
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
ResetRoleInheritance|xsd:boolean|Declares whether to reset the role inheritance or not for the site, if it is a sub-site, optional attribute.
CopyRoleAssignments|xsd:boolean|Defines whether to copy role assignments or not while breaking role inheritance, optional attribute.
RemoveExistingUniqueRoleAssignments|xsd:boolean|Defines whether to remove unique role assignments or not if the site already breaks role inheritance. If true all existing unique role assignments on the site will be removed if BreakRoleInheritance also is true.
ClearSubscopes|xsd:boolean|Defines whether to clear subscopes or not while breaking role inheritance for the site, optional attribute.
AssociatedGroups|ReplaceableString|Specifies the list of groups that are associated with the Web site. Groups in this list will appear under the Groups section in the People and Groups page.
AssociatedOwnerGroup|ReplaceableString|Specifies the default owners group for this site. The group will automatically be added to the end of the Associated Groups list.
AssociatedMemberGroup|ReplaceableString|Specifies the default members group for this site. The group will automatically be added to the top of the Associated Groups list.
AssociatedVisitorGroup|ReplaceableString|Specifies the default visitors group for this site. The group will automatically be added to the end of the Associated Groups list.
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
   <pnp:SearchNavigation />
</pnp:Navigation>
```


Here follow the available child elements for the Navigation element.


Element|Type|Description
-------|----|-----------
GlobalNavigation|[GlobalNavigation](#globalnavigation)|The Global Navigation settings for the Provisioning Template, optional element.
CurrentNavigation|[CurrentNavigation](#currentnavigation)|The Current Navigation settings for the Provisioning Template, optional element.
SearchNavigation|[StructuralNavigation](#structuralnavigation)|The Search Navigation settings for the Provisioning Template, optional element.

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
      IsExternal="xsd:boolean">
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
<pnp:UsersList
      ClearExistingItems="xsd:boolean">
   <pnp:User />
</pnp:UsersList>
```


Here follow the available child elements for the UsersList element.


Element|Type|Description
-------|----|-----------
User|[User](#user)|

Here follow the available attributes for the UsersList element.


Attibute|Type|Description
--------|----|-----------
ClearExistingItems|xsd:boolean|Declares whether to clear existing users before adding new users, optional attribute.
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
      ValidationMessage="xsd:string"
      TemplateInternalName="xsd:string">
   <pnp:PropertyBagEntries />
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
PropertyBagEntries|[PropertyBagEntries](#propertybagentries)|The Property Bag entries of the root folder, optional collection of elements.
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
TemplateInternalName|xsd:string|Defines the alternate template internal name for a list based on a .STP file/list definition.
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
      Name="xsd:string"
      ContentTypeID="pnp:ContentTypeId">
   <pnp:Folder />
   <pnp:Security />
   <pnp:PropertyBagEntries />
   <pnp:DefaultColumnValues />
   <pnp:Properties />
</pnp:Folder>
```


Here follow the available child elements for the Folder element.


Element|Type|Description
-------|----|-----------
Folder|[Folder](#folder)|A child Folder of another Folder item, optional element.
Security|[ObjectSecurity](#objectsecurity)|Defines the security rules for the row that will be added to the List Instance, optional element.
PropertyBagEntries|[PropertyBagEntries](#propertybagentries)|The Property Bag entries of the Folder, optional collection of elements.
DefaultColumnValues|[DefaultColumnValues](#defaultcolumnvalues)|The Default Columne Values entries of the Folder, optional collection of elements.
Properties|[FileProperties](#fileproperties)|The File Properties for the Folder, optional collection of elements.

Here follow the available attributes for the Folder element.


Attibute|Type|Description
--------|----|-----------
Name|xsd:string|The Name of the Folder, required attribute.
ContentTypeID|ContentTypeId|The Content Type ID for the Folder, optional attribute.
<a name="defaultcolumnvalues"></a>
### DefaultColumnValues
The Default Columne Values entries of the Folder, optional collection of elements.

```xml
<pnp:DefaultColumnValues>
   <pnp:DefaultColumnValue />
</pnp:DefaultColumnValues>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
DefaultColumnValue|[StringDictionaryItem](#stringdictionaryitem)|A custom Default Column Value for the Provisioning Template, collection of elements.
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
<pnp:ContentTypeFieldRef
      UpdateChildren="xsd:boolean">
</pnp:ContentTypeFieldRef>
```


Here follow the available attributes for the ContentTypeFieldRef element.


Attibute|Type|Description
--------|----|-----------
UpdateChildren|xsd:boolean|Declares whether the current field reference has to be udpated on inherited content types, optional attribute.
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
      ExpiresInDays="pnp:ReplaceableInt"
      ClientState="pnp:ReplaceableString">
</pnp:Webhook>
```


Here follow the available attributes for the Webhook element.


Attibute|Type|Description
--------|----|-----------
ServerNotificationUrl|ReplaceableString|The Server Notification URL of the Webhook, required attribute.
ExpiresInDays|ReplaceableInt|The expire days for the subscription of the Webhook, required attribute.
ClientState|ReplaceableString|An opaque string passed back to the client on all notifications, optional attribute.
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
<a name="baseclientsidepage"></a>
### BaseClientSidePage
Defines a base Client Side Page.

```xml
<pnp:BaseClientSidePage
      PromoteAsNewsArticle="xsd:boolean"
      PromoteAsTemplate="xsd:boolean"
      Overwrite="xsd:boolean"
      Layout="xsd:string"
      Publish="xsd:boolean"
      EnableComments="xsd:boolean"
      Title="xsd:string"
      ContentTypeID="pnp:ReplaceableString"
      ThumbnailUrl="pnp:ReplaceableString">
   <pnp:Header />
   <pnp:Sections />
   <pnp:FieldValues />
   <pnp:Security />
   <pnp:Properties />
</pnp:BaseClientSidePage>
```


Here follow the available child elements for the BaseClientSidePage element.


Element|Type|Description
-------|----|-----------
Header|[Header](#header)|Defines the layout of the Header for the current client side page
Sections|[Sections](#sections)|Defines the Canvas sections for a single ClientSidePage.
FieldValues|[FieldValues](#fieldvalues)|Defines the page fields values, if any.
Security|[ObjectSecurity](#objectsecurity)|
Properties|[Properties](#properties)|Defines property bag properties for the client side page, optional element.

Here follow the available attributes for the BaseClientSidePage element.


Attibute|Type|Description
--------|----|-----------
PromoteAsNewsArticle|xsd:boolean|Declares to promote the page as a news article.
PromoteAsTemplate|xsd:boolean|Declares to promote the page as a page template.
Overwrite|xsd:boolean|Can the page be overwritten if it exists.
Layout|xsd:string|Defines the target layout for the client-side page, optional attribute (default: Article).
Publish|xsd:boolean|Defines whether the page will be published or not, optional attribute (default: true).
EnableComments|xsd:boolean|Defines whether the page will have comments enabled or not, optional attribute (default: true).
Title|xsd:string|Defines the Title of the page, optional attribute.
ContentTypeID|ReplaceableString|Defines the Content Type ID for the page, optional attribute.
ThumbnailUrl|ReplaceableString|Defines the URL of the thumbnail for the client side page, optional attribute.
<a name="header"></a>
### Header
Defines the layout of the Header for the current client side page

```xml
<pnp:Header
      Type=""
      ServerRelativeImageUrl="xsd:string"
      TranslateX="xsd:double"
      TranslateY="xsd:double"
      LayoutType=""
      TextAlignment=""
      ShowTopicHeader="xsd:boolean"
      ShowPublishDate="xsd:boolean"
      TopicHeader="xsd:string"
      AlternativeText="xsd:string"
      Authors="xsd:string"
      AuthorByLine="xsd:string"
      AuthorByLineId="xsd:int">
</pnp:Header>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Type||Defines the layout of the Header for the current client side page
ServerRelativeImageUrl|xsd:string|Defines the server-relative URL of the image for the header of the current client side page.
TranslateX|xsd:double|Defines the x-translate of the image for the header of the current client side page.
TranslateY|xsd:double|Defines the y-translate of of the image for the header of the current client side page.
LayoutType||Defines the type of layout used inside the header of the current client side page
TextAlignment||Defines the text alignment of the text in the header of the current client side page
ShowTopicHeader|xsd:boolean|Defines whether to show the topic header in the title region of the current client side page.
ShowPublishDate|xsd:boolean|Defines whether to show the page publication date in the title region of the current client side page.
TopicHeader|xsd:string|Defines the topic header text to show if ShowTopicHeader is set to true of the current client side page.
AlternativeText|xsd:string|Defines the alternative text for the header image of the current client side page.
Authors|xsd:string|Defines the page author(s) to be displayed of the current client side page.
AuthorByLine|xsd:string|Defines the page author by line of the current client side page.
AuthorByLineId|xsd:int|Defines the ID of the page author by line of the current client side page.
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
<a name="fieldvalues"></a>
### FieldValues
Defines the page fields values, if any.

```xml
<pnp:FieldValues>
   <pnp:FieldValue />
</pnp:FieldValues>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
FieldValue|[StringDictionaryItem](#stringdictionaryitem)|A custom field for the current Client Page
<a name="properties"></a>
### Properties
Defines property bag properties for the client side page, optional element.

```xml
<pnp:Properties>
   <pnp:Property />
</pnp:Properties>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Property|[StringDictionaryItem](#stringdictionaryitem)|A custom property for the current Client Page
<a name="clientsidepage"></a>
### ClientSidePage
Defines a single element of type ClientSidePage.

```xml
<pnp:ClientSidePage
      PageName="xsd:string"
      LCID="xsd:int"
      CreateTranslations="xsd:boolean">
</pnp:ClientSidePage>
```


Here follow the available attributes for the ClientSidePage element.


Attibute|Type|Description
--------|----|-----------
PageName|xsd:string|Defines the page name for a single ClientSidePage.
LCID|xsd:int|The Locale ID of a Localization Language, optional attribute.
CreateTranslations|xsd:boolean|Defines whether to create translations of the current Client Side Page, optional attribute.
<a name="translatedclientsidepage"></a>
### TranslatedClientSidePage


```xml
<pnp:TranslatedClientSidePage
      LCID="xsd:int"
      PageName="xsd:string">
</pnp:TranslatedClientSidePage>
```


Here follow the available attributes for the TranslatedClientSidePage element.


Attibute|Type|Description
--------|----|-----------
LCID|xsd:int|The Locale ID of a Localization Language, optional attribute.
PageName|xsd:string|Defines the page name for a single ClientSidePage, optional attribute.
<a name="header"></a>
### Header
Defines the Header settings for the target site.

```xml
<pnp:Header
      Layout=""
      MenuStyle=""
      BackgroundEmphasis="pnp:Emphasis">
</pnp:Header>
```


Here follow the available attributes for the Header element.


Attibute|Type|Description
--------|----|-----------
Layout||Defines the Layout of the Header, required attribute.
MenuStyle||Defines the Menu Style, required attribute.
BackgroundEmphasis|Emphasis|Defines the Background Emphasis of the Header, optional attribute.
<a name="footer"></a>
### Footer
Defines the Footer settings for the target site.

```xml
<pnp:Footer
      Enabled="xsd:boolean"
      Logo="pnp:ReplaceableString"
      Name="pnp:ReplaceableString"
      RemoveExistingNodes="xsd:boolean">
   <pnp:FooterLinks />
</pnp:Footer>
```


Here follow the available child elements for the Footer element.


Element|Type|Description
-------|----|-----------
FooterLinks|[FooterLinks](#footerlinks)|Defines the Footer Links for the target site.

Here follow the available attributes for the Footer element.


Attibute|Type|Description
--------|----|-----------
Enabled|xsd:boolean|Defines whether the site Footer is enabled or not, required attribute.
Logo|ReplaceableString|Defines the Logo to render in the Footer, optional attribute.
Name|ReplaceableString|Defines the name of the footer, optional attribute.
RemoveExistingNodes|xsd:boolean|Defines whether the existing site Footer links should be removed, optional attribute.
<a name="footerlinks"></a>
### FooterLinks
Defines the Footer Links for the target site.

```xml
<pnp:FooterLinks>
   <pnp:FooterLink />
</pnp:FooterLinks>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
FooterLink|[FooterLink](#footerlink)|Defines a Footer Link for the target site.
<a name="footerlink"></a>
### FooterLink
Defines a Footer Link for the target site.

```xml
<pnp:FooterLink
      DisplayName="pnp:ReplaceableString"
      Url="pnp:ReplaceableString">
   <pnp:FooterLink />
</pnp:FooterLink>
```


Here follow the available child elements for the FooterLink element.


Element|Type|Description
-------|----|-----------
FooterLink|[FooterLink](#footerlink)|Defines a collection of children Footer Link for the current Footer Link (which represents an header).

Here follow the available attributes for the FooterLink element.


Attibute|Type|Description
--------|----|-----------
DisplayName|ReplaceableString|Defines the DisplayName for the Footer Link for the target site.
Url|ReplaceableString|Defines the URL for the Footer Link for the target site.
<a name="provisioningwebhook"></a>
### ProvisioningWebhook


```xml
<pnp:ProvisioningWebhook
      Kind=""
      Url="pnp:ReplaceableString"
      Method=""
      BodyFormat=""
      Async="xsd:boolean">
   <pnp:Parameters />
</pnp:ProvisioningWebhook>
```


Here follow the available child elements for the ProvisioningWebhook element.


Element|Type|Description
-------|----|-----------
Parameters|[Parameters](#parameters)|A collection of custom parameters that will be provided to the webhook request

Here follow the available attributes for the ProvisioningWebhook element.


Attibute|Type|Description
--------|----|-----------
Kind||Defines the kind of a Provisioning Template Webhook, required attribute.
Url|ReplaceableString|Defines the URL of a Provisioning Template Webhook, can be a replaceable string and it is a required attribute.
Method||Defines how to call the target Webhook URL, required attribute.
BodyFormat||Defines how to format the request body for HTTP POST requests, optional attribute.
Async|xsd:boolean|Defines whether the Provisioning Template Webhook should be executed asychronously or not, optional attribute.
<a name="parameters"></a>
### Parameters
A collection of custom parameters that will be provided to the webhook request

```xml
<pnp:Parameters>
   <pnp:Parameter />
</pnp:Parameters>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Parameter|[StringDictionaryItem](#stringdictionaryitem)|A custom parameter that will be provided to the webhook request
<a name="canvassection"></a>
### CanvasSection
A Canvas Section for a Client-side Page.

```xml
<pnp:CanvasSection
      Order="xsd:int"
      Type=""
      BackgroundEmphasis="pnp:Emphasis"
      VerticalSectionEmphasis="pnp:Emphasis">
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
BackgroundEmphasis|Emphasis|The emphasis color of the Canvas Section for a Client-side Page.
VerticalSectionEmphasis|Emphasis|The emphasis color of the Canvas Section for a Client-side Page.
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
      Enabled="xsd:boolean"
      NoDefaultOrigins="xsd:boolean">
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
NoDefaultOrigins|xsd:boolean|Defines whether the CDN should have default origins, optional attribute.
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
      ID="xsd:ID">
   <pnp:SiteCollections />
   <pnp:TermStore />
</pnp:Sequence>
```


Here follow the available child elements for the Sequence element.


Element|Type|Description
-------|----|-----------
SiteCollections|[SiteCollections](#sitecollections)|A collection of Site Collection elements to provision through a Sequence, optional element.
TermStore|[TermStore](#termstore)|A Term Store to provision through a Sequence, optional element.

Here follow the available attributes for the Sequence element.


Attibute|Type|Description
--------|----|-----------
ID|xsd:ID|A unique identifier of the Sequence, required attribute.
<a name="sitecollections"></a>
### SiteCollections
A collection of Site Collection elements to provision through a Sequence, optional element.

```xml
<pnp:SiteCollections>
   <pnp:SiteCollection />
</pnp:SiteCollections>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
SiteCollection|[SiteCollection](#sitecollection)|A Site Collection to provision through a Sequence, optional element.
<a name="sitecollection"></a>
### SiteCollection
Defines the base element for a SiteCollection that will be created into the target tenant/farm.

```xml
<pnp:SiteCollection>
   <pnp:Templates />
   <pnp:Sites />
</pnp:SiteCollection>
```


Here follow the available child elements for the SiteCollection element.


Element|Type|Description
-------|----|-----------
Templates|[Templates](#templates)|Templates that can be provisioned together with the Site Collection, optional collection of elements.
Sites|[Sites](#sites)|Allows to create sub-Sites under the current site
<a name="templates"></a>
### Templates
Templates that can be provisioned together with the Site Collection, optional collection of elements.

```xml
<pnp:Templates>
   <pnp:ProvisioningTemplateReference />
</pnp:Templates>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
ProvisioningTemplateReference|[ProvisioningTemplateReference](#provisioningtemplatereference)|Defines a reference to a Provisioning Template defined or imported in the current Provisioning File
<a name="sites"></a>
### Sites
Allows to create sub-Sites under the current site

```xml
<pnp:Sites>
   <pnp:Site />
</pnp:Sites>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Site|[Site](#site)|A Site to provision through a Sequence, optional element.
<a name="communicationsite"></a>
### CommunicationSite


```xml
<pnp:CommunicationSite
      Url="pnp:ReplaceableString"
      Owner="pnp:ReplaceableString"
      SiteDesign="pnp:ReplaceableString"
      AllowFileSharingForGuestUsers="xsd:boolean"
      Classification="pnp:ReplaceableString"
      Language="pnp:ReplaceableString">
</pnp:CommunicationSite>
```


Here follow the available attributes for the CommunicationSite element.


Attibute|Type|Description
--------|----|-----------
Url|ReplaceableString|The URL of the target Site, required attribute.
Owner|ReplaceableString|Primary Owner of the target Site, required attribute.
SiteDesign|ReplaceableString|The ID of the SiteDesign, if any, to apply to the target Site, optional attribute.
AllowFileSharingForGuestUsers|xsd:boolean|Defines whether the target Site can be shared to guest users or not, optional attribute.
Classification|ReplaceableString|The Classification of the target Site, if any, optional attribute.
Language|ReplaceableString|Language of the target Site, optional attribute.
<a name="teamsite"></a>
### TeamSite


```xml
<pnp:TeamSite
      Alias="pnp:ReplaceableString"
      DisplayName="pnp:ReplaceableString"
      IsPublic="xsd:boolean"
      Classification="pnp:ReplaceableString"
      Teamify="xsd:boolean"
      HideTeamify="xsd:boolean"
      GroupLifecyclePolicyId="pnp:ReplaceableString"
      Language="pnp:ReplaceableString">
</pnp:TeamSite>
```


Here follow the available attributes for the TeamSite element.


Attibute|Type|Description
--------|----|-----------
Alias|ReplaceableString|The Alias of the target Site, required attribute.
DisplayName|ReplaceableString|The DisplayName of the target Site, required attribute.
IsPublic|xsd:boolean|Defines whether the Office 365 Group for the target Site is Public or Private, required attribute.
Classification|ReplaceableString|The Classification of the target Site, if any, optional attribute.
Teamify|xsd:boolean|Defines whether to create a Microsoft Team backing the modern Team Site, optional attribute.
HideTeamify|xsd:boolean|Defines whether to hide the create a Microsoft Team option in the UI of the Team Site, optional attribute.
GroupLifecyclePolicyId|ReplaceableString|Allows to associate the Office 365 Group associated with the Team Site to a Group Lifecycle Policy, optional attribute.
Language|ReplaceableString|Language of the target Site, optional attribute.
<a name="teamsitenogroup"></a>
### TeamSiteNoGroup


```xml
<pnp:TeamSiteNoGroup
      Url="pnp:ReplaceableString"
      Owner="pnp:ReplaceableString"
      TimeZoneId="pnp:ReplaceableString"
      Language="pnp:ReplaceableString">
</pnp:TeamSiteNoGroup>
```


Here follow the available attributes for the TeamSiteNoGroup element.


Attibute|Type|Description
--------|----|-----------
Url|ReplaceableString|The URL of the target Site, required attribute.
Owner|ReplaceableString|Primary Owner of the target Site, required attribute.
TimeZoneId|ReplaceableString|TimeZone of the target Site, optional attribute.
Language|ReplaceableString|Language of the target Site, optional attribute.
<a name="site"></a>
### Site
Defines a Site that will be created into a target Site Collection.

```xml
<pnp:Site>
   <pnp:Sites />
   <pnp:Templates />
</pnp:Site>
```


Here follow the available child elements for the Site element.


Element|Type|Description
-------|----|-----------
Sites|[Sites](#sites)|Allows to create sub-Sites under the current site
Templates|[Templates](#templates)|Templates that can be provisioned together with the Site Collection, optional collection of elements.
<a name="sites"></a>
### Sites
Allows to create sub-Sites under the current site

```xml
<pnp:Sites>
   <pnp:Site />
</pnp:Sites>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Site|[Site](#site)|A Site to provision through a Sequence, optional element.
<a name="templates"></a>
### Templates
Templates that can be provisioned together with the Site Collection, optional collection of elements.

```xml
<pnp:Templates>
   <pnp:ProvisioningTemplateReference />
</pnp:Templates>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
ProvisioningTemplateReference|[ProvisioningTemplateReference](#provisioningtemplatereference)|Defines a reference to a Provisioning Template defined or imported in the current Provisioning File
<a name="teamsubsitenogroup"></a>
### TeamSubSiteNoGroup


```xml
<pnp:TeamSubSiteNoGroup
      Url="pnp:ReplaceableString"
      TimeZoneId="pnp:ReplaceableString"
      Language="pnp:ReplaceableString">
</pnp:TeamSubSiteNoGroup>
```


Here follow the available attributes for the TeamSubSiteNoGroup element.


Attibute|Type|Description
--------|----|-----------
Url|ReplaceableString|The URL of the target Site, required attribute.
TimeZoneId|ReplaceableString|TimeZone of the target Site, optional attribute.
Language|ReplaceableString|Language of the target Site, optional attribute.
<a name="termstore"></a>
### TermStore
A TermStore to use for provisioning of TermGroups.

```xml
<pnp:TermStore>
   <pnp:TermGroup />
</pnp:TermStore>
```


Here follow the available child elements for the TermStore element.


Element|Type|Description
-------|----|-----------
TermGroup|[TermGroup](#termgroup)|The TermGroup element to provision into the target TermStore through, optional collection of elements.
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
      Owner="pnp:ReplaceableString"
      Description="xsd:string"
      IsAvailableForTagging="xsd:boolean">
</pnp:TermSetItem>
```


Here follow the available attributes for the TermSetItem element.


Attibute|Type|Description
--------|----|-----------
Owner|ReplaceableString|The Owner of the Term Set Item, optional attribute.
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
      SourceTermId="pnp:GUID"
      ReuseChildren="xsd:boolean"
      IsPinned="xsd:boolean"
      IsPinnedRoot="xsd:boolean">
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
ReuseChildren|xsd:boolean|Declares if children of the source term should be reused, optional attribute.
IsPinned|xsd:boolean|Declares if this term is pinned, optional attribute. Requires IsReused set to true.
IsPinnedRoot|xsd:boolean|Declares if this term is the root of the pinned hierarchy, optional attribute. Requires IsPinned and IsReused set to true.
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
<a name="teams"></a>
### Teams
Entry point to manage Microsoft Teams provisioning.

```xml
<pnp:Teams>
   <pnp:TeamTemplate />
   <pnp:Team />
   <pnp:Apps />
</pnp:Teams>
```


Here follow the available child elements for the Teams element.


Element|Type|Description
-------|----|-----------
TeamTemplate|[TeamTemplate](#teamtemplate)|Defines a new Team to create starting from a JSON template.
Team|[TeamWithSettings](#teamwithsettings)|Defines a new Team to create or update.
Apps|[TeamsApps](#teamsapps)|Defines the Apps for Teams
<a name="baseteam"></a>
### BaseTeam
Defines the base information for a Team of Microsoft Teams.

```xml
<pnp:BaseTeam
      DisplayName="pnp:ReplaceableString"
      Description="pnp:ReplaceableString"
      Classification="pnp:ReplaceableString"
      Visibility=""
      Photo="pnp:ReplaceableString">
</pnp:BaseTeam>
```


Here follow the available attributes for the BaseTeam element.


Attibute|Type|Description
--------|----|-----------
DisplayName|ReplaceableString|The Display Name of the Team, required attribute.
Description|ReplaceableString|The Description of the Team, required attribute.
Classification|ReplaceableString|The Classification for the Team, optional attribute.
Visibility||The Visibility for the Team, optional attribute.
Photo|ReplaceableString|Declares the Photo for the Team, optional attribute.
<a name="teamtemplate"></a>
### TeamTemplate


```xml
<pnp:TeamTemplate>
</pnp:TeamTemplate>
```

<a name="teamwithsettings"></a>
### TeamWithSettings
Defines a new Team to create or update

```xml
<pnp:TeamWithSettings
      GroupId="pnp:ReplaceableString"
      TargetSiteUrl="pnp:ReplaceableString"
      ProvisioningTemplateId="pnp:ReplaceableString"
      Specialization=""
      CloneFrom="pnp:ReplaceableString"
      Archived="xsd:boolean"
      MailNickname="pnp:ReplaceableString">
</pnp:TeamWithSettings>
```


Here follow the available attributes for the TeamWithSettings element.


Attibute|Type|Description
--------|----|-----------
GroupId|ReplaceableString|Declares the ID of the targt Group/Team to update, optional attribute. Cannot be used together with CloneFrom or with ProvisioningTemplateId.
TargetSiteUrl|ReplaceableString|Declares the URL of the targt Site to teamify, optional attribute. Cannot be used together with CloneFrom or with ProvisioningTemplateId.
ProvisioningTemplateId|ReplaceableString|Declares the ID of a Provisioning Template to apply to the SharePoint Online site backing the current Team, optional attribute. Cannot be used together with CloneFrom or with TargetSiteUrl.
Specialization||The Specialization for the Team, optional attribute.
CloneFrom|ReplaceableString|Declares the ID of another Team to Clone the current Team from, optional attribute. Cannot be used together with GroupId, or with TargetSiteUrl.
Archived|xsd:boolean|Declares whether the Team is archived or not, optional attribute.
MailNickname|ReplaceableString|Declares the nickname for the Team, optional attribute. Required for new Teams.
<a name="teamsecurity"></a>
### TeamSecurity
Defines the Security settings for the Team, optional element.

```xml
<pnp:TeamSecurity
      AllowToAddGuests="xsd:boolean">
   <pnp:Owners />
   <pnp:Members />
</pnp:TeamSecurity>
```


Here follow the available child elements for the TeamSecurity element.


Element|Type|Description
-------|----|-----------
Owners|[TeamSecurityUsers](#teamsecurityusers)|Defines the Owners of the Team, optional element.
Members|[TeamSecurityUsers](#teamsecurityusers)|Defines the Members of the Team, optional element.

Here follow the available attributes for the TeamSecurity element.


Attibute|Type|Description
--------|----|-----------
AllowToAddGuests|xsd:boolean|Defines whether guests are allowed in the Team, optional attribute.
<a name="teamsecurityusers"></a>
### TeamSecurityUsers
Defines a list of users for a the Team, optional element.

```xml
<pnp:TeamSecurityUsers
      ClearExistingItems="xsd:boolean">
   <pnp:User />
</pnp:TeamSecurityUsers>
```


Here follow the available child elements for the TeamSecurityUsers element.


Element|Type|Description
-------|----|-----------
User|[User](#user)|Defines a user for a the Team, optional element.

Here follow the available attributes for the TeamSecurityUsers element.


Attibute|Type|Description
--------|----|-----------
ClearExistingItems|xsd:boolean|Declares whether to clear existing users before adding new users, optional attribute.
<a name="user"></a>
### User
Defines a user for a the Team, optional element.

```xml
<pnp:User
      UserPrincipalName="pnp:ReplaceableString">
</pnp:User>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
UserPrincipalName|ReplaceableString|Defines User Principal Name (UPN) of the target user, required attribute.
<a name="teamchannel"></a>
### TeamChannel
Defines a Channel for a Team, optional element.

```xml
<pnp:TeamChannel
      DisplayName="pnp:ReplaceableString"
      Description="pnp:ReplaceableString"
      IsFavoriteByDefault="xsd:boolean"
      ID="pnp:ReplaceableString">
   <pnp:Tabs />
   <pnp:TabResources />
   <pnp:Messages />
</pnp:TeamChannel>
```


Here follow the available child elements for the TeamChannel element.


Element|Type|Description
-------|----|-----------
Tabs|[TeamChannelTabs](#teamchanneltabs)|Defines a collection of Tabs for a Channel in a Team, optional element.
TabResources|[TeamTabResources](#teamtabresources)|Defines a collection of Resources for Tabs in a Team Channel, optional element.
Messages|[TeamChannelMessages](#teamchannelmessages)|Defines a collection of Messages for a Team Channel, optional element.

Here follow the available attributes for the TeamChannel element.


Attibute|Type|Description
--------|----|-----------
DisplayName|ReplaceableString|Defines the Display Name of the Channel, required attribute.
Description|ReplaceableString|Defines the Description of the Channel, required attribute.
IsFavoriteByDefault|xsd:boolean|Defines whether the Channel is Favorite by default for all members of the Team, optional attribute.
ID|ReplaceableString|Declares the ID for the Channel, optional attribute.
<a name="teamchanneltabs"></a>
### TeamChannelTabs
Defines a collection of Tabs for a Channel in a Team, optional element.

```xml
<pnp:TeamChannelTabs>
   <pnp:Tab />
</pnp:TeamChannelTabs>
```


Here follow the available child elements for the TeamChannelTabs element.


Element|Type|Description
-------|----|-----------
Tab|[Tab](#tab)|Defines a Tab for a Channel in a Team, optional element.
<a name="tab"></a>
### Tab
Defines a Tab for a Channel in a Team, optional element.

```xml
<pnp:Tab
      DisplayName="pnp:ReplaceableString"
      TeamsAppId="pnp:ReplaceableString"
      ID="pnp:ReplaceableString">
   <pnp:Configuration />
</pnp:Tab>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
Configuration|[Configuration](#configuration)|Defines the Configuration for the Tab, required element.

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
DisplayName|ReplaceableString|Defines the Display Name of the Channel, required attribute.
TeamsAppId|ReplaceableString|App definition identifier of the tab, required attribute.
ID|ReplaceableString|Declares the ID for the Tab, optional attribute.
<a name="teamtabresources"></a>
### TeamTabResources
Defines a collection of Resources for Tabs of a Channel in a Team.

```xml
<pnp:TeamTabResources>
   <pnp:TabResource />
</pnp:TeamTabResources>
```


Here follow the available child elements for the TeamTabResources element.


Element|Type|Description
-------|----|-----------
TabResource|[TabResource](#tabresource)|Defines a Resource for a Tab in a Channel of a Team.
<a name="tabresource"></a>
### TabResource
Defines a Resource for a Tab in a Channel of a Team.

```xml
<pnp:TabResource
      Type=""
      TargetTabId="pnp:ReplaceableString">
   <pnp:TabResourceSettings />
</pnp:TabResource>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
TabResourceSettings|[ReplaceableString](#replaceablestring)|Defines the Configuration for the Tab Resource, optional element.

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
Type||Defines the Type of Resource for the Tab, required attribute.
TargetTabId|ReplaceableString|Defines the ID of the target Tab for the Resource, required attribute.
<a name="teamchannelmessages"></a>
### TeamChannelMessages
Defines a collection of Messages for a Channel in a Team, optional element.

```xml
<pnp:TeamChannelMessages>
   <pnp:Message />
</pnp:TeamChannelMessages>
```


Here follow the available child elements for the TeamChannelMessages element.


Element|Type|Description
-------|----|-----------
Message|xsd:string|Defines a Message for a Channel in a Team, optional element.
<a name="teamsapps"></a>
### TeamsApps
Defines the Apps for Teams

```xml
<pnp:TeamsApps>
   <pnp:App />
</pnp:TeamsApps>
```


Here follow the available child elements for the TeamsApps element.


Element|Type|Description
-------|----|-----------
App|[App](#app)|Defines a Teams App to add or update.
<a name="app"></a>
### App
Defines a Teams App to add or update.

```xml
<pnp:App
      AppId="xsd:string"
      PackageUrl="pnp:ReplaceableString">
</pnp:App>
```


Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
AppId|xsd:string|Unique ID - from PnP perspective - for the App, defined for further reference in the Provisioning Template, required attribute.
PackageUrl|ReplaceableString|The URL or path for the Teams App package, required attribute.
<a name="azureactivedirectory"></a>
### AzureActiveDirectory
Entry point to manage Microsoft Azure Active Directory provisioning.

```xml
<pnp:AzureActiveDirectory>
   <pnp:Users />
</pnp:AzureActiveDirectory>
```


Here follow the available child elements for the AzureActiveDirectory element.


Element|Type|Description
-------|----|-----------
Users|[AADUsers](#aadusers)|Defines a list of users that will be created in the target tenant
<a name="aadusers"></a>
### AADUsers
Defines a list of users that will be created in the target tenant

```xml
<pnp:AADUsers>
   <pnp:User />
</pnp:AADUsers>
```


Here follow the available child elements for the AADUsers element.


Element|Type|Description
-------|----|-----------
User|[User](#user)|
<a name="user"></a>
### User


```xml
<pnp:User
      AccountEnabled="xsd:boolean"
      DisplayName="pnp:ReplaceableString"
      MailNickname="pnp:ReplaceableString"
      PasswordPolicies="pnp:ReplaceableString"
      UserPrincipalName="pnp:ReplaceableString"
      ProfilePhoto="pnp:ReplaceableString"
      GivenName="pnp:ReplaceableString"
      Surname="pnp:ReplaceableString"
      JobTitle="pnp:ReplaceableString"
      OfficeLocation="pnp:ReplaceableString"
      PreferredLanguage="pnp:ReplaceableString"
      MobilePhone="pnp:ReplaceableString"
      UsageLocation="pnp:ReplaceableString">
   <pnp:PasswordProfile />
   <pnp:Licenses />
</pnp:User>
```


Here follow the available child elements for the  element.


Element|Type|Description
-------|----|-----------
PasswordProfile|[PasswordProfile](#passwordprofile)|The Password Profile for the user, required element.
Licenses|[Licenses](#licenses)|Defines a collection of licenses to activate/associate with the user, optional element.

Here follow the available attributes for the  element.


Attibute|Type|Description
--------|----|-----------
AccountEnabled|xsd:boolean|Declares whether the user's account is enabled or not, required attribute.
DisplayName|ReplaceableString|The Display Name of the user, required attribute.
MailNickname|ReplaceableString|The Mail Nickname of the user, required attribute.
PasswordPolicies|ReplaceableString|The Password Policies for the user, optional attribute.
UserPrincipalName|ReplaceableString|The UPN for the user, required attribute.
ProfilePhoto|ReplaceableString|The URL of the Photo for the user, optional attribute.
GivenName|ReplaceableString|Declares the GivenName of the user, optional attribute.
Surname|ReplaceableString|Declares the Surname of the user, optional attribute.
JobTitle|ReplaceableString|Declares the Job Title of the user, optional attribute.
OfficeLocation|ReplaceableString|Declares the Office Location of the user, optional attribute.
PreferredLanguage|ReplaceableString|Declares the Preferred Language of the user, optional attribute.
MobilePhone|ReplaceableString|Declares the Mobile Phone of the user, optional attribute.
UsageLocation|ReplaceableString|Declares the Usage Location for licenses assigned to the user, required attribute.
<a name="drive"></a>
### Drive
Entry point to manage OneDrive for Business provisioning.

```xml
<pnp:Drive>
   <pnp:DriveRoot />
</pnp:Drive>
```


Here follow the available child elements for the Drive element.


Element|Type|Description
-------|----|-----------
DriveRoot|[DriveRoot](#driveroot)|Defines a list of Drives in the target OneDrive for Business
<a name="driveroot"></a>
### DriveRoot
Entry point to manage OneDrive for Business Drives provisioning.

```xml
<pnp:DriveRoot
      DriveUrl="pnp:ReplaceableString">
   <pnp:DriveItems />
</pnp:DriveRoot>
```


Here follow the available child elements for the DriveRoot element.


Element|Type|Description
-------|----|-----------
DriveItems|[DriveItems](#driveitems)|Defines a list of DriveItems to be uploaded/updated in the target OneDrive for Business drive

Here follow the available attributes for the DriveRoot element.


Attibute|Type|Description
--------|----|-----------
DriveUrl|ReplaceableString|Defines the relative URL of the target Drive item in OneDrive for Business. The target Drive must exist and will not be created by the Provisioning Engine.
<a name="driveitems"></a>
### DriveItems
Entry point to define a collection of DriveItems in the OneDrive for Business provisioning.

```xml
<pnp:DriveItems>
   <pnp:DriveFolder />
   <pnp:DriveFile />
</pnp:DriveItems>
```


Here follow the available child elements for the DriveItems element.


Element|Type|Description
-------|----|-----------
DriveFolder|[DriveFolder](#drivefolder)|Defines DriveItem to be uploaded/updated in the target OneDrive for Business drive
DriveFile|[DriveFile](#drivefile)|Defines DriveItem to be uploaded/updated in the target OneDrive for Business drive
<a name="drivefolder"></a>
### DriveFolder
Entry point to define a Folder in the OneDrive for Business provisioning.

```xml
<pnp:DriveFolder
      Name="pnp:ReplaceableString"
      Src="pnp:ReplaceableString"
      Overwrite="xsd:boolean"
      Recursive="xsd:boolean"
      IncludedExtensions="xsd:string"
      ExcludedExtensions="xsd:string">
   <pnp:DriveFolder />
   <pnp:DriveFile />
</pnp:DriveFolder>
```


Here follow the available child elements for the DriveFolder element.


Element|Type|Description
-------|----|-----------
DriveFolder|[DriveFolder](#drivefolder)|Defines DriveItem to be uploaded/updated in the target OneDrive for Business drive
DriveFile|[DriveFile](#drivefile)|Defines DriveItem to be uploaded/updated in the target OneDrive for Business drive

Here follow the available attributes for the DriveFolder element.


Attibute|Type|Description
--------|----|-----------
Name|ReplaceableString|Defines the Name of the Folder in OneDrive for Business.
Src|ReplaceableString|Defines the Source path of the folder in OneDrive for Business. If provided, the whole folder will be uploaded.
Overwrite|xsd:boolean|The Overwrite flag for the File items in the Directory, optional attribute.
Recursive|xsd:boolean|Defines whether to recursively browse through all the child folders of the source Folder, optional attribute.
IncludedExtensions|xsd:string|The file Extensions (lower case) to include while uploading the source Folder, optional attribute.
ExcludedExtensions|xsd:string|The file Extensions (lower case) to exclude while uploading the source Folder, optional attribute.
<a name="drivefile"></a>
### DriveFile
Entry point to define a DriveItem in the OneDrive for Business provisioning.

```xml
<pnp:DriveFile
      Name="pnp:ReplaceableString"
      Src="pnp:ReplaceableString"
      Overwrite="xsd:boolean">
</pnp:DriveFile>
```


Here follow the available attributes for the DriveFile element.


Attibute|Type|Description
--------|----|-----------
Name|ReplaceableString|Defines the Name of the DriveItem in OneDrive for Business.
Src|ReplaceableString|Defines the Source path of the file in OneDrive for Business.
Overwrite|xsd:boolean|The Overwrite flag for the File items in the Directory, optional attribute.
<a name="provisioningwebhooks"></a>
### ProvisioningWebhooks
Allows to define one or more webhooks that will be invoked by the engine, while applying the provisioning

```xml
<pnp:ProvisioningWebhooks>
   <pnp:ProvisioningWebhook />
</pnp:ProvisioningWebhooks>
```


Here follow the available child elements for the ProvisioningWebhooks element.


Element|Type|Description
-------|----|-----------
ProvisioningWebhook|[ProvisioningWebhook](#provisioningwebhook)|Defines a single Provisioning Webhook
