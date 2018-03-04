Office 365 Developer PnP Core Component Provisioning Engine Tokens
==================================================================

### Summary ###
The Office 365 Developer PnP Core Provisioning Engine supports certain tokens which will be replaced by corresponding values during provisioning.
These tokens can be used to make the template site collection independent for instance.

Below all the supported tokens are listed:


|Token|Example|Output example|Description
|-----|-------|-----------|-----
|{associatedmembergroup}|{associatedmembergroup}|Members|Will return the name of the site's associated members group.|
|{associatedownergroup}|{associatedownergroup}|Owners|Will return the name of the site's associated owners group.|
|{associatedvisitorgroup}|{associatedvisitorgroup}|Vistors|Will return the name of the site's associated visitors group.|
|{contenttypeid:&lt;name&gt;}|{contenttypeid:Reservations}|0x0102004F51EFDEA49C49668EF9C6744C8CF87D|Will return the id of the content type by name.|
|{currentuserfullname}|{currentuserfullname}|Test User|Will return the full name of the user. Notice, does not work when using app only authentication.|
|{currentuserid}|{currentuserid}|12|Will return the current user id, as present in the Site User Info List|
|{currentuserloginname}|{currentuserloginname}|i:0#.f\|membership\|user@domain.com|Returns the current login name of the user. Notice that this does not work when using app only authentication|
|{datenow}|{datenow}|2017-01-13T22:53:15.908Z|Returns the current date and time converted to UTC and formatted as "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffK"|
|{fieldtitle:&lt;internalname&gt;}|{fieldtitle:Title}|Title|Returns the title (displayname) of a field by its internal name|
|{groupid:&lt;name&gt;}|{groupid:Site Members}|5|Returns the id of the group by name|
|{guid}|{guid}|678149a4-208e-491e-9e93-a2b0d754f5e4|Returns a new guid without { }. Write {{guid}} to return a guid in the shape of {678149a4-208e-491e-9e93-a2b0d754f5e4}|
|{keywordstermstoreid}|{keywordstermstoreid}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212E|Will return the ID/Guid of the keyword term store, without { }. If you want a ID with { } around the value, use the token as follows: {{keywordstermstoreid}}|
|{listid:&lt;name&gt;}|{listid:Demo List}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212E|Will return the ID of the list specified by the parameter, which is the title of the list. If you want a ID with { } around the value, use the token as follows: {{listid:Demo List}}|
|{listurl:&lt;name&gt;}|{listurl:Demo List}|lists/demolist|Will return the url of the list specified by the parameter, which is the title of the list.|
|{viewid:&lt;ListName&gt;,&lt;ViewName&gt;}|{viewid:MyList,MyView}|ddc063cb-0c0e-4ce1-821c-a1f629992e42|Returns the id of a given view in a list without { }. Write {{viewid:MyList,MyView}} to return an id in the shape of {ddc063cb-0c0e-4ce1-821c-a1f629992e42}|
|{loc:&lt;token&gt;}<br/>{localize:&lt;token&gt;}<br/>{localization:&lt;token&gt;}<br/>{resource:&lt;token&gt;}<br/>{res:&lt;token&gt;}|{resource:MyListName}|Mijn lijst|Returns a token from an embedded resource file in a template for the current language of the web you are provisioning to.|
|{masterpagecatalog}|{masterpagecatalog}|/sites/demo/_catalogs/masterpage|Will return the server relative url of the masterpage catalog for the current site.|
|{parameter:&lt;name&gt;}|{parameter:DefaultGroup}|string value|Will return the value of the parameter as specified in the template.|
|{roledefinition:&lt;name&gt;}|{roledefinition:Administrator}|Object of type RoleDefinition|Returns a roledefinition, to be used in Security elements of the template. Eg. Administrator, Contributor, Reader|
|{sitecollectionid}|{sitecollectionid}|73170a53-c1ce-4cd0-9569-e464069f1a69|Returns the id of the current site collection. Write {{sitecollectionid}} to return the id in the shape of {73170a53-c1ce-4cd0-9569-e464069f1a69}|
|{sitecollectiontermgroupid}|{sitecollectiontermgroupid}|2235e428-83a9-4486-9583-64dd454f9918|Returns the id of the current site collection term group. This group is not present by default in SharePoint 2013 and 2016, but if the token is encountered in the template and the group does not exist, it will be created.|
|{sitecollectiontermgroupname}|{sitecollectiontermgroupname}|Site Collection - mytenant.sharepoint.com-sites-demo|Returns the name of the site collection term group.  You can use this value also in as a nested token, alike {termsetid:{sitecollectiontermgroupname}:mytermset}|
|{sitecollectiontermsetid:&lt;name&gt;}|{sitecollectiontermsetid:Departments}|52a3abcd-4dec-4b9a-b5ba-f9220f8d47bd|Returns the id of a specific termset in the site collection term group.|
|{sitecollectiontermstoreid}|{sitecollectiontermstoreid}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212E|Will return the ID/Guid of the site collection term store without enclosing { }. If you want a ID with { } around the value, use the token as follows: {{sitecollectiontermstoreid}}.|
|{sitecollection}|{sitecollection}|/sites/demo|Will return the server relative URL of the current site collection rootweb|
|{siteid}|{siteid}|cb779dae-0b29-4cec-b3ac-9983d3389ad0|Returns the id of the current web.|
|{sitename}|{sitename}|My Demo Site|Returns the title of the current web.|
|{sitetitle}|{sitetitle}|My Demo Site|Returns the title of the current web.|
|{siteowner}|{siteowner}|i:0#.f\|membership\|user@domain.com|Returns the login name of the current owner of the site.|
|{site}|{site}|/sites/demo/test|Will returm the server relative URL of the current web.|
|{termsetid:&lt;Group&gt;:&lt;Set&gt;}|{termsetid:TestGroup:TestSet}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212|Will return the ID of the termset that is residing under the specified group. If you want a ID with { } around the value, use the token as follows: {{termsetid:TestGroup:TestSet}}.|
|{termstoreid:&lt;name&gt;}|{termstoreid:ExtraStore}|d42bcad2-0603-4b86-8e3d-72177f4519ca|Returns the id of a termstore by its name.|
|{themecatalog}|{themecatalog}|/sites/demo/_catalogs/theme|Will return the server relative url of the current site theme catalog.|
|{webpartid:&lt;name&gt;}|{webpartid:MyWebPart}|767245f6-5f47-4cb5-b558-bcc04956bb7b|Returns the id of a webpart by its name.|
|{webname}|{webname}|testsite|Returns the Name property value of the current web. The name part is the last part of the URL, e.g. given a web with /sites/testsite/subweb1/subweb2 this token will return 'subweb2'| 
<img src="https://telemetry.sharepointpnp.com/pnp-sites-core/core/provisioningenginetokens" /> 
