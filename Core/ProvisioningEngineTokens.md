Office 365 Developer PnP Core Component Provisioning Engine Tokens
==================================================================

### Summary ###
The Office 365 Developer PnP Core Provisioning Engine supports certain tokens which will be replaced by corresponding values during provisioning.
These tokens can be used to make the template site collection independent for instance.

Below all the supported tokens are listed:

|Token|Example|Output example|Description
|-----|-------|-----------|-----
|{keywordtermstoreid}|{keywordtermstoreid}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212E|Will return the ID/Guid of the keyword term store, without { }. If you want a ID with { } around the value, use the token as follows: {{keywordtermstoreid}}|
|{listid:&lt;name&gt;}|{listid:Demo List}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212E|Will return the ID of the list specified by the parameter, which is the title of the list. If you want a ID with { } around the value, use the token as follows: {{listid:Demo List}}|
|{listurl:&lt;name&gt;}|{listurl:Demo List}|lists/demolist|Will return the url of the list specified by the parameter, which is the title of the list.|
|{masterpagecatalog}|{masterpagecatalog}|/sites/demo/_catalogs/masterpage|Will return the server relative url of the masterpage catalog for the current site.|
|{parameter:&lt;name&gt;}|{parameter:DefaultGroup}|string value|Will return the value of the parameter as specified in the template.|
|{sitecollectiontermstoreid}|{sitecollectiontermstoreid}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212E|Will return the ID/Guid of the site collection term store without enclosing { }. If you want a ID with { } around the value, use the token as follows: {{sitecollectiontermstoreid}}.|
|{sitecollection}|{sitecollection}|/sites/demo|Will return the server relative URL of the current site collection rootweb|
|{site}|{site}|/sites/demo/test|Will returm the server relative URL of the current web.|
|{termsetid:&lt;Group&gt;:&lt;Set&gt;}|{termsetid:TestGroup:TestSet}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212|Will return the ID of the termset that is residing under the specified group. If you want a ID with { } around the value, use the token as follows: {{termsetid:TestGroup:TestSet}}.|
|{themecatalog}|{themecatalog}|/sites/demo/_catalogs/theme|Will return the server relative url of the current site theme catalog.|
|{sitecollectiontermgroupid}|{sitecollectiontermgroupid}|FDF19D89-A82F-4AB9-9BB5-B49E6CA5212|Will return the ID of the default site collection termgroup. This is group is not present by default, but if the token is encountered in a template and the group does not exist, it will be created.|
|{sitecollectiontermgroupname}|{sitecollectiontermgroupname}|Site Collection - erwinmcm.sharepoint.com-sites-projects|Will return the name of the default site collection termgroupo. This is group is not present by default, but if the token is encountered in a template and the group does not exist, it will be created. You can use this value also in as a nested token, alike {termsetid:{sitecollectiontermgroupname}:mytermset}|
|{currentuserid}|{currentuserid}|12|Will return the current user id, as present in the Site User Info List|
|{currentuserfullname}|{currentuserfullname}|Test User|Will return the full name of the user|
|{currentuserloginname}|{currentuserloginname}|i:0#.f&#124;membership&#124;testuser@domain.com|Will return the login name of the user|
|{guid}|{guid}|f17b152f-029c-4288-b4c8-d11101f05cb1|Will return a new guid|