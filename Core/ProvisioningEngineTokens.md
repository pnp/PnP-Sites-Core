Office 365 Developer PnP Core Component Provisioning Engine Tokens
==================================================================

### Summary ###
The SharePoint PnP Core Provisioning Engine supports certain tokens which will be replaced by corresponding values during provisioning. These tokens can be used to make the template site collection independent for instance.

Below all the supported tokens are listed:

Token|Description|Example|Returns
:-----|:----------|:------|:------
{apppackageid:[packagename]}|Returns the ID of an app package given its name|{apppackageid:MyPackageName}|55898e77-a7bf-4799-8034-506db5521b98
{associatedmembergroup}|Returns the title of the associated members SharePoint group of a site|{associatedmembergroup}|My Site Members Group Title
{associatedownergroup}|Returns the title of the associated owners SharePoint group of a site|{associatedownergroup}|My Site Owners Group Title
{associatedvisitorgroup}|Returns the title of the associated visitors SharePoint group of a site|{associatedvisitorgroup}|My Site Visitors Group Title
{authenticationrealm}|Returns the authentication ID of the current tenant/farm|{authenticationrealm}|55898e77-a7bf-4799-8034-506db5521b98
{contenttypeid:[contenttypename]}|Returns the ID of the specified content type|{contenttypeid:My Content Type}|0x0102004F51EFDEA49C49668EF9C6744C8CF87D
{currentuserfullname}|Returns the full name of the current user e.g. the user using the engine.|{currentuserfullname}|John Doe
{currentuserid}|Returns the ID of the current user e.g. the user using the engine.|{currentuserid}|4
{currentuserloginname}|Returns the login name of the current user e.g. the user using the engine.|{currentuserloginname}|i:0#.f|membership|user@domain.com
{fieldtitle:[internalname]}|Returns the title/displayname of a field given its internalname|{fieldtitle:LeaveEarly}|Leaving Early
{fileuniqueid:[siteRelativePath]}|Returns the unique id of a file which is being provisioned by the current template.|{fileuniqueid:/sitepages/home.aspx}|f2cd6d5b-1391-480e-a3dc-7f7f96137382
{fileuniqueidencoded:[siteRelativePath]}|Returns the html safe encoded unique id of a file which is being provisioned by the current template.|{fileuniqueid:/sitepages/home.aspx}|f2cd6d5b%2D1391%2D480e%2Da3dc%2D7f7f96137382
{groupid:[groupname]}|Returns the id of a SharePoint group given its name|{groupid:My Site Owners}|6
{guid}|Returns a newly generated GUID|{guid}|f2cd6d5b-1391-480e-a3dc-7f7f96137382
{hosturl}|Returns a full url of the current host|{hosturl}|https://mycompany.sharepoint.com
{keywordstermstoreid}|Returns a id of the default keywords term store|{keywordstermstoreid}|f2cd6d5b-1391-480e-a3dc-7f7f96137382
{listid:[name]}|Returns a id of the list given its name|{listid:My List}|f2cd6d5b-1391-480e-a3dc-7f7f96137382
{listurl:[name]}|Returns a site relative url of the list given its name|{listid:My List}|Lists/MyList
{localization:[key]}|Returns a value from a in the template provided resource file given the locale of the site that the template is applied to|{localization:MyListTitle}|My List Title
{masterpagecatalog}|Returns a server relative url of the master page catalog|{masterpagecatalog}|/sites/mysite/_catalogs/masterpage
{now}|Returns the current date in universal date time format: yyyy-MM-ddTHH:mm:ss.fffK|{now}|2018-04-18T15:44:45.898+02:00
{pageuniqueid:[siterelativepath]}|Returns the id of a client side page that is being provisioned through the current template|{pageuniqueid:SitePages/Home.aspx}|767bc144-e605-4d8c-885a-3a980feb39c6
{pageuniqueidencoded:[siterelativepath]}|Returns the HTML safe encoded id of a client side page that is being provisioned through the current template|{pageuniqueidencoded:SitePages/Home.aspx}|767bc144%2De605%2D4d8c%2D885a%2D3a980feb39c6
{parameter:[parametername]}|Returns the value of a parameter defined in the template|{parameter:MyParameter}|the value of the parameter
{realm}|Returns the authentication ID of the current tenant/farm|{realm}|55898e77-a7bf-4799-8034-506db5521b98
{roledefinition:[roletype]}|Returns the name of role definition given the role type|{roledefinition:Editor}|Editors
{roledefinitionid:[rolename]}|Returns the id of the given role definition name|{roledefinitionid:My Role Definition}|23
{sequencesiteurl:[provisioningid]}|Returns a full url of the site given its provisioning ID from the sequence|{sequencesiteurl:MYID}|https://contoso.sharepoint.com/sites/mynewsite
{site}|Returns the server relative url of the current site|{site}|/sites/mysitecollection/mysite
{sitecollection}|Returns the server relative url of the site collection|{sitecollection}|/sites/mysitecollection
{sitecollectionconnectedoffice365groupid}|Returns the ID of the Office 365 group connected to the current site|{sitecollectionconnectedoffice365groupid}|767bc144-e605-4d8c-885a-3a980feb39c6
{sitecollectionidencoded}|Returns the HTML safe id of the site collection|{sitecollectionidencoded}|767bc144%2De605%2D4d8c%2D885a%2D3a980feb39c6
{sitecollectionidencoded}|Returns the id of the site collection|{sitecollectionidencoded}|767bc144-e605-4d8c-885a-3a980feb39c6
{sitecollectiontermgroupid}|Returns the id of the site collection term group|{sitecollectiontermgroupid}|767bc144-e605-4d8c-885a-3a980feb39c6
{sitecollectiontermgroupname}|Returns the name of the site collection term group|{sitecollectiontermgroupname}|Site Collection - mytenant.sharepoint.com-sites-mysite
{sitecollectiontermsetid:[termsetname]}|Returns the id of the given termset name located in the sitecollection termgroup|{sitecollectiontermsetid:MyTermset}|9188a794-cfcf-48b6-9ac5-df2048e8aa5d
{sitecollectiontermstoreid}|Returns the id of the given default site collection term store|{sitecollectiontermstoreid}|9188a794-cfcf-48b6-9ac5-df2048e8aa5d
{sitedesignid:[designtitle]}|Returns the id of the given site design|{sitedesignid:My Site Design}|9188a794-cfcf-48b6-9ac5-df2048e8aa5d
{siteid}|Returns the id of the current site|{siteid}|9188a794-cfcf-48b6-9ac5-df2048e8aa5d
{siteidencoded}|Returns the id of the current site|{siteidencoded}|9188a794%2Dcfcf%2D48b6%2D9ac5%2Ddf2048e8aa5d
{sitename}|Returns the title of the current site|{sitename}|My Company Portal
{siteowner}|Returns the login name of the current site owner|{siteowner}|i:0#.f|membership|user@domain.com
{sitescriptid:[scripttitle]}|Returns the id of the given site script|{sitescriptid:My Site Script}|9188a794-cfcf-48b6-9ac5-df2048e8aa5d
{storageentityvalue:[key]}|Returns the value of a storage entity provided by the key|{storageentityvalue:MyKey}|My Value
{termsetid:[groupname]:[termsetname]}|Returns the id of a term set given its name and its parent group|{termsetid:MyGroup:MyTermset}|9188a794-cfcf-48b6-9ac5-df2048e8aa5d
{termstoreid:[storename]}|Returns the id of a term store given its name|{termstoreid:MyTermStore}|9188a794-cfcf-48b6-9ac5-df2048e8aa5d
{themecatalog}|Returns the server relative url of the theme catalog|{themecatalog}|/sites/sitecollection/_catalogs/theme
{viewid:[listname],[viewname]}|Returns a id of the view given its name for a given list|{viewid:My List,My View}|f2cd6d5b-1391-480e-a3dc-7f7f96137382
{webname}|Returns the name part of the URL of the Server Relative URL of the Web|{webname}|MyWeb
{webpartid:[webpartname]}|Returns the id of a webpart that is being provisioned to a page through a template|{webpartid:mywebpart}|66e2b037-f749-402d-90b2-afd643850c26
