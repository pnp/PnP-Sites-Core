# PnP Unit Test report for OnlineAppOnly on Friday, February 19, 2016 #
This page is showing the results of the PnP unit test run.

## Test configuration ##
This report contains the unit test results from the following run:

Parameter | Value
----------|------
PnP Unit Test configuration | OnlineAppOnly
Test run date | Friday, February 19, 2016
Test run time | 11:35 PM
PnP branch | dev
Visual Studio build configuration | debug

## Test summary ##
During this test run 294 tests have been executed with following outcome:

Parameter | Value
----------|------
Executed tests | 294
Elapsed time | 0h 32m 29s
Passed tests | 114
Failed tests | **144**
Skipped tests | 36
Was canceled | False
Was aborted | False
Error | 

## Test run details ##

### Failed tests ###
<table>
<tr>
<td><b>Test name</b></td>
<td><b>Test outcome</b></td>
<td><b>Duration</b></td>
<td><b>Message</b></td>
</tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CanAddContentTypeToListByName</td><td>Failed</td><td>0h 0m 22s</td><td>TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CanRemoveContentTypeFromListByName</td><td>Failed</td><td>0h 0m 46s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CanRemoveContentTypeFromListByName threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CanRemoveContentTypeFromListById</td><td>Failed</td><td>0h 0m 48s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CanRemoveContentTypeFromListById threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CreateExistingFieldTest</td><td>Failed</td><td>0h 0m 46s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CreateExistingFieldTest threw exception System.Net.WebException, but exception System.ArgumentException was expected. Exception message: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.GetContentTypeByIdTest</td><td>Failed</td><td>0h 0m 45s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.GetContentTypeByIdTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.RemoveFieldByInternalNameThrowsOnNoMatchTest</td><td>Failed</td><td>0h 0m 47s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.RemoveFieldByInternalNameThrowsOnNoMatchTest threw exception System.Net.WebException, but exception System.ArgumentException was expected. Exception message: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CreateFieldFromXmlTest</td><td>Failed</td><td>0h 0m 47s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CreateFieldFromXmlTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByNameTest</td><td>Failed</td><td>0h 0m 45s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByNameTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByIdTest</td><td>Failed</td><td>0h 0m 47s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByIdTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByNameInSubWebTest</td><td>Failed</td><td>0h 0m 3s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByNameInSubWebTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByIdInSubWebTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByIdInSubWebTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByNameSearchInSiteHierarchyTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByNameSearchInSiteHierarchyTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByIdSearchInSiteHierarchyTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ContentTypeExistsByIdSearchInSiteHierarchyTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.AddFieldToContentTypeTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.AddFieldToContentTypeTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.AddFieldToContentTypeMakeRequiredTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.AddFieldToContentTypeMakeRequiredTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.SetDefaultContentTypeToListTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.SetDefaultContentTypeToListTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ReorderContentTypesTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.ReorderContentTypesTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CreateContentTypeByXmlTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CreateContentTypeByXmlTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkToWebTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkToWebTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkToSiteTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkToSiteTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkIEnumerableToWebTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkIEnumerableToWebTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkIEnumerableToSiteTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsLinkIEnumerableToSiteTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.DeleteJsLinkFromWebTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.DeleteJsLinkFromWebTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.DeleteJsLinkFromSiteTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.DeleteJsLinkFromSiteTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsBlockToWebTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsBlockToWebTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsBlockToSiteTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.JavaScriptExtensionsTests.AddJsBlockToSiteTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.ListExtensionsTests.CreateListTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method Microsoft.SharePoint.Client.Tests.ListExtensionsTests.CreateListTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.ListRatingExtensionTest.EnableRatingExperienceTest</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method Microsoft.SharePoint.Client.Tests.ListRatingExtensionTest.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.ListRatingExtensionTest.EnableLikesExperienceTest</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method Microsoft.SharePoint.Client.Tests.ListRatingExtensionTest.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.GetAdministratorsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddAdministratorsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddGroupTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.GroupExistsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddPermissionLevelToGroupTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddPermissionLevelToGroupSubSiteTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddPermissionLevelToGroupListTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddPermissionLevelToGroupListItemTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.RemovePermissionLevelFromGroupSubSiteTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddPermissionLevelByRoleDefToGroupTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddPermissionLevelToUserTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddPermissionLevelToUserTestByRoleDefTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddSamePermissionLevelTwiceToGroupTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddReaderAccessToEveryoneExceptExternalsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.AddReaderAccessToEveryoneTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.GetAllUniqueRoleAssignmentsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.SecurityExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.StructuralNavigationExtensionsTests.GetNavigationSettingsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.StructuralNavigationExtensionsTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.StructuralNavigationExtensionsTests.UpdateNavigationSettingsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.StructuralNavigationExtensionsTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.StructuralNavigationExtensionsTests.UpdateNavigationSettings2Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.StructuralNavigationExtensionsTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.CheckOutFileTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.CheckInFileTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.UploadFileTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.UploadFileWebDavTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.VerifyIfUploadRequiredTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.SetFilePropertiesTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.GetFileTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.EnsureSiteFolderTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.EnsureLibraryFolderTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.FileFolderExtensionsTests.EnsureLibraryFolderRecursiveTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.FileFolderExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.GetSiteCollectionsTest</td><td>Failed</td><td>0h 0m 34s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.GetOneDriveSiteCollectionsTest</td><td>Failed</td><td>0h 0m 33s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.GetUserProfileServiceClientTest</td><td>Failed</td><td>0h 0m 32s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.CheckIfSiteExistsTest</td><td>Failed</td><td>0h 0m 35s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.SiteExistsTest</td><td>Failed</td><td>0h 0m 35s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.SubSiteExistsTest</td><td>Failed</td><td>0h 0m 34s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.CreateDeleteSiteCollectionTest</td><td>Failed</td><td>0h 0m 36s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.TenantExtensionsTests.SetSiteLockStateTest</td><td>Failed</td><td>0h 0m 36s</td><td>Initialization method OfficeDevPnP.Core.Tests.AppModelExtensions.TenantExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline-admin.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.NavigationExtensionsTests.AddTopNavigationNodeTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.NavigationExtensionsTests.AddTopNavigationNodeTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.NavigationExtensionsTests.AddQuickLaunchNodeTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.NavigationExtensionsTests.AddQuickLaunchNodeTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.NavigationExtensionsTests.AddSearchNavigationNodeTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.NavigationExtensionsTests.AddSearchNavigationNodeTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.NavigationExtensionsTests.DeleteTopNavigationNodeTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.NavigationExtensionsTests.DeleteTopNavigationNodeTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.NavigationExtensionsTests.DeleteQuickLaunchNodeTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.NavigationExtensionsTests.DeleteQuickLaunchNodeTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.NavigationExtensionsTests.DeleteSearchNavigationNodeTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.NavigationExtensionsTests.DeleteSearchNavigationNodeTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.NavigationExtensionsTests.DeleteAllNavigationNodesTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.NavigationExtensionsTests.DeleteAllNavigationNodesTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorGetFile1Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorGetFile2Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorGetFiles1Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorGetFiles2Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorGetFiles3Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorGetFileBytes1Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorGetFileBytes2Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorSaveStream1Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorSaveStream2Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorSaveStream3Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorDelete1Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorSharePointTests.SharePointConnectorDelete2Test</td><td>Failed</td><td>0h 0m 0s</td><td>Class Initialization method OfficeDevPnP.Core.Tests.Framework.Connectors.ConnectorSharePointTests.ClassInit threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderReceivesExpectedParameters</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderReceivesExpectedParameters threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderProvidesTokens</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderProvidesTokens threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderCanBeDisabled</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderCanBeDisabled threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectAuditSettingsTests.CanExtractAuditSettings</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectAuditSettingsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectAuditSettingsTests.CanProvisionAuditSettings</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectAuditSettingsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectRegionalSettingsTests.CanExtractRegionalSettings</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectRegionalSettingsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectRegionalSettingsTests.CanProvisionRegionalSettings</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectRegionalSettingsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectSupportedUILanguagesTests.CanExtractSupportedUILanguages</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectSupportedUILanguagesTests.CanExtractSupportedUILanguages threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectSupportedUILanguagesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectSupportedUILanguagesTests.CanProvisionSupportedUILanguages</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectSupportedUILanguagesTests.CanProvisionSupportedUILanguages threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectSupportedUILanguagesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectComposedLookTests.CanCreateComposedLooks</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectComposedLookTests.CanCreateComposedLooks threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectPagesTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPagesTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPagesTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectPagesTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPagesTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPagesTests.Cleanup threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectSiteSecurityTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectSiteSecurityTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectSiteSecurityTests.CanCreateEntities1</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectSiteSecurityTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectSiteSecurityTests.CanCreateEntities2</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectSiteSecurityTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectPropertyBagEntryTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPropertyBagEntryTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPropertyBagEntryTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectPropertyBagEntryTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPropertyBagEntryTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectPropertyBagEntryTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectListInstanceTests.FolderContentTypeShouldNotBeRemovedFromProvisionedDocumentLibraries</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.FolderContentTypeShouldNotBeRemovedFromProvisionedDocumentLibraries threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectListInstanceTests.UpdatedListTitleShouldBeAvailableAsToken</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.UpdatedListTitleShouldBeAvailableAsToken threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectListInstanceTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFilesTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFilesTests.CanAddWebPartsToForms</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CanAddWebPartsToForms threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFilesTests.CanProvisionObjectsRequiredField</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CanProvisionObjectsRequiredField threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFilesTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFilesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFeaturesTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFeaturesTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFeaturesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFeaturesTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFeaturesTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFeaturesTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectCustomActionsTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectCustomActionsTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectCustomActionsTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectCustomActionsTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectCustomActionsTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectCustomActionsTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFieldTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFieldTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFieldTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectFieldTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFieldTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectFieldTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectContentTypeTests.CanProvisionObjects</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectContentTypeTests.CanProvisionObjects threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectContentTypeTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectContentTypeTests.CanCreateEntities</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectContentTypeTests.CanCreateEntities threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'
TestCleanup method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.ObjectContentTypeTests.CleanUp threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.TokenParserTests.ParseTests</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ObjectHandlers.TokenParserTests.ParseTests threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.Providers.BaseTemplateTests.GetBaseTemplateForCurrentSiteTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.Providers.BaseTemplateTests.GetBaseTemplateForCurrentSiteTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeToXml</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeToXml threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.GetRemoteTemplateTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.Framework.ProvisioningTemplates.DomainModelTests.GetRemoteTemplateTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.SetPropertyBagValueIntTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.SetPropertyBagValueStringTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.SetPropertyBagValueMultipleRunsTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.RemovePropertyBagValueTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.GetPropertyBagValueIntTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.GetPropertyBagValueStringTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.PropertyBagContainsKeyTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.GetIndexedPropertyBagKeysTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.AddIndexedPropertyBagKeyTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.RemoveIndexedPropertyBagKeyTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.GetProvisioningTemplateTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.GetAppInstancesTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.RemoveAppInstanceByTitleTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.IsSubWebTest</td><td>Failed</td><td>0h 0m 0s</td><td>Initialization method Microsoft.SharePoint.Client.Tests.WebExtensionsTests.Initialize threw exception. System.Net.WebException: System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'.</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.CanAddLayoutToWikiPageTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.CanAddLayoutToWikiPageTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.CanAddHtmlToWikiPageTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.CanAddHtmlToWikiPageTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.ProveThatWeCanAddHtmlToPageAfterChangingLayoutTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.ProveThatWeCanAddHtmlToPageAfterChangingLayoutTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.CanCreatePublishingPageTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.CanCreatePublishingPageTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.PublishingPageWithInvalidCharsIsCorrectlyCreatedTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.PublishingPageWithInvalidCharsIsCorrectlyCreatedTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.CanCreatePublishedPublishingPageWhenModerationIsEnabledTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.CanCreatePublishedPublishingPageWhenModerationIsEnabledTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.CanCreatePublishedPublishingPageWhenModerationIsDisabledTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.CanCreatePublishedPublishingPageWhenModerationIsDisabledTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>
<tr><td>Tests.AppModelExtensions.PageExtensionsTests.CreatedPublishingPagesSetsTitleCorrectlyTest</td><td>Failed</td><td>0h 0m 0s</td><td>Test method OfficeDevPnP.Core.Tests.AppModelExtensions.PageExtensionsTests.CreatedPublishingPagesSetsTitleCorrectlyTest threw exception: 
System.Net.WebException: The remote name could not be resolved: 'bertonline.sharepoint.com'</td></tr>

</table>


### Skipped tests ###
<table>
<tr>
<td><b>Test name</b></td>
<td><b>Test outcome</b></td>
<td><b>Duration</b></td>
<td><b>Message</b></td>
</tr>
<tr><td>Microsoft.SharePoint.Client.Tests.ListExtensionsTests.SetDefaultColumnValuesTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Tests.AppModelExtensions.FileChangeTrackingTests.OOBMasterPagesHaveChangedTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. No test automation SQL database information found...or found database is not reachable.</td></tr>
<tr><td>Tests.AppModelExtensions.SearchExtensionsTests.SetSiteCollectionSearchCenterUrlTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Search tests are not supported when testing using app-only</td></tr>
<tr><td>Tests.AppModelExtensions.SearchExtensionsTests.GetSearchConfigurationFromWebTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Search tests are not supported when testing using app-only</td></tr>
<tr><td>Tests.AppModelExtensions.SearchExtensionsTests.GetSearchConfigurationFromSiteTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Search tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.CreateTaxonomyFieldTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.CreateTaxonomyFieldMultiValueTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.SetTaxonomyFieldValueTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.CreateTaxonomyFieldLinkedToTermSetTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.CreateTaxonomyFieldLinkedToTermTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetTaxonomySessionTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetDefaultKeywordsTermStoreTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetDefaultSiteCollectionTermStoreTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetTermSetsByNameTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetTermGroupByNameTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetTermGroupByIdTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetTermByNameTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.GetTaxonomyItemByPathTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.AddTermToTermsetTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.AddTermToTermsetWithTermIdTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ImportTermsTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ImportTermsToTermStoreTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ImportTermSetSampleShouldCreateSetTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ImportTermSetShouldUpdateSetTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ImportTermSetShouldUpdateByGuidTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ExportTermSetTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ExportTermSetFromTermstoreTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.TaxonomyExtensionsTests.ExportAllTermsTest</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectTermGroupsTests.CanProvisionToSiteCollectionTermGroupUsingToken</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Tests.Framework.ObjectHandlers.ObjectTermGroupsTests.CanProvisionObjects</td><td>Skipped</td><td>0h 0m 0s</td><td>Assert.Inconclusive failed. Taxonomy tests are not supported when testing using app-only</td></tr>
<tr><td>Tests.Framework.Providers.BaseTemplateTests.ExtractSingleTemplate2</td><td>Skipped</td><td>0h 0m 0s</td><td></td></tr>
<tr><td>Tests.Framework.Providers.BaseTemplateTests.ExtractBaseTemplates2</td><td>Skipped</td><td>0h 0m 0s</td><td></td></tr>
<tr><td>Tests.Framework.Providers.BaseTemplateTests.DumpBaseTemplates</td><td>Skipped</td><td>0h 0m 0s</td><td></td></tr>
<tr><td>Tests.Framework.Providers.BaseTemplateTests.DumpSingleTemplate</td><td>Skipped</td><td>0h 0m 0s</td><td></td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.InstallSolutionTest</td><td>Skipped</td><td>0h 0m 0s</td><td></td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.WebExtensionsTests.UninstallSolutionTest</td><td>Skipped</td><td>0h 0m 0s</td><td></td></tr>

</table>


### Passed tests ###
<table>
<tr>
<td><b>Test name</b></td>
<td><b>Test outcome</b></td>
<td><b>Duration</b></td>
</tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.CanUploadHtmlPageLayoutAndConvertItToAspxVersionTest</td><td>Passed</td><td>0h 1m 29s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.CanUploadPageLayoutTest</td><td>Passed</td><td>0h 1m 13s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.CanUploadPageLayoutWithPathTest</td><td>Passed</td><td>0h 1m 10s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.AllowAllPageLayoutsTest</td><td>Passed</td><td>0h 1m 8s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.DeployThemeAndCreateComposedLookTest</td><td>Passed</td><td>0h 1m 9s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.ComposedLookExistsTest</td><td>Passed</td><td>0h 1m 9s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.GetCurrentComposedLookTest</td><td>Passed</td><td>0h 1m 33s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.CreateComposedLookShouldWorkTest</td><td>Passed</td><td>0h 1m 4s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.CreateComposedLookByNameShouldWorkTest</td><td>Passed</td><td>0h 1m 5s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.SetComposedLookInheritsTest</td><td>Passed</td><td>0h 1m 55s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.SetComposedLookResetInheritanceTest</td><td>Passed</td><td>0h 2m 35s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.SeattleMasterPageIsUnchangedTest</td><td>Passed</td><td>0h 1m 6s</td></tr>
<tr><td>Tests.AppModelExtensions.BrandingExtensionsTests.IsSubsiteTest</td><td>Passed</td><td>0h 1m 7s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.NotLoadedPropertyExceptionTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsurePropertyTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.NotLoadedCollectionExceptionTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsureCollectionPropertyTest</td><td>Passed</td><td>0h 0m 1s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.NotLoadedComplexPropertyExceptionTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsureComplexPropertyTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsureMultiplePropertiesTest</td><td>Passed</td><td>0h 0m 1s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsurePropertiesIncludeTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsurePropertyIncludeTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsureComplexPropertyWithDependencyTest</td><td>Passed</td><td>0h 0m 1s</td></tr>
<tr><td>Tests.AppModelExtensions.ClientObjectExtensionsTests.EnsureComplexPropertiesWithDependencyTest</td><td>Passed</td><td>0h 0m 1s</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FeatureExtensionsTests.ActivateSiteFeatureTest</td><td>Passed</td><td>0h 0m 1s</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FeatureExtensionsTests.ActivateWebFeatureTest</td><td>Passed</td><td>0h 0m 2s</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FeatureExtensionsTests.DeactivateSiteFeatureTest</td><td>Passed</td><td>0h 0m 1s</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FeatureExtensionsTests.DeactivateWebFeatureTest</td><td>Passed</td><td>0h 0m 1s</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FeatureExtensionsTests.IsSiteFeatureActiveTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FeatureExtensionsTests.IsWebFeatureActiveTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Microsoft.SharePoint.Client.Tests.FieldAndContentTypeExtensionsTests.CreateFieldTest</td><td>Passed</td><td>0h 0m 6s</td></tr>
<tr><td>Tests.Diagnostics.LogTests.LogTest1</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorGetFile1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorGetFile2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorGetFiles1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorGetFiles2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorGetFileBytes1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorGetFileBytes2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorSaveStream1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorSaveStream2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorSaveStream3Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorDelete1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorAzureTests.AzureConnectorDelete2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorGetFile1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorGetFile2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorGetFile3Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorGetFiles1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorGetFiles2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorGetFileBytes1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorSaveStream1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorSaveStream2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorSaveStream3Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorDelete1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Connectors.ConnectorFileSystemTests.FileConnectorDelete2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.CanProviderCallOut</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.ProviderCallOutThrowsException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.ProviderAssemblyMissingThrowsAgrumentException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.ProviderTypeNameMissingThrowsAgrumentException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.ProviderClientCtxIsNullThrowsAgrumentNullException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderCallOutThrowsException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderAssemblyMissingThrowsAgrumentException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderTypeNameMissingThrowsAgrumentException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ExtensibilityCallOut.ExtensibilityTests.TokenProviderClientCtxIsNullThrowsAgrumentNullException</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLFileSystemGetTemplatesTest</td><td>Passed</td><td>0h 0m 2s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLFileSystemGetTemplate1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLFileSystemGetTemplate2Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLAzureStorageGetTemplatesTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLAzureStorageGetTemplate1Test</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLAzureStorageGetTemplate2SecureTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLFileSystemConvertTemplatesFromV201503toV201505</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.ResolveSchemaFormatV201503</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.ResolveSchemaFormatV201505</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLResolveValidXInclude</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.Providers.XMLProvidersTests.XMLResolveInvalidXInclude</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanDeserializeXMLToDomainObject1</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectToXML1</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectToXMLStream1</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanDeserializeXMLToDomainObject2</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectToXML2</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetTemplateNameandVersion</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetPropertyBagEntries</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetOwners</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetAdministrators</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetMembers</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetVistors</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetFeatures</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanGetCustomActions</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeToJSon</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.ValidateFullProvisioningSchema5</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.ValidateSharePointProvisioningSchema6</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanDeserializeXMLToDomainObject5</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanDeserializeXMLToDomainObject6</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectToXML6</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectToXML5ByIdentifier</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectToXML5ByFileLink</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectWithJsonFormatter</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanHandleDomainObjectWithJsonFormatter</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanDeserializeXMLToDomainObjectFrom201512Full</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.CanSerializeDomainObjectToXML201512Full</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Framework.ProvisioningTemplates.DomainModelTests.AreTemplatesEqual</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Utilities.Tests.JsonUtilityTests.SerializeTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Utilities.Tests.JsonUtilityTests.DeserializeTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Utilities.Tests.JsonUtilityTests.DeserializeListTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Utilities.Tests.JsonUtilityTests.DeserializeListIsNotFixedSizeTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Utilities.Tests.JsonUtilityTests.DeserializeListNoDataStillWorksTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Utilities.EncryptionUtilityTests.ToSecureStringTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Utilities.EncryptionUtilityTests.ToInSecureStringTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Utilities.EncryptionUtilityTests.EncryptStringWithDPAPITest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Utilities.EncryptionUtilityTests.DecryptStringWithDPAPITest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.Diagnostics.PnPMonitoredScopeTests.PnPMonitoredScopeNestingTest</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.UrlUtilityTests.ContainsInvalidCharsReturnsFalseForValidString</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.UrlUtilityTests.ContainsInvalidUrlCharsReturnsTrueForInvalidString</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.UrlUtilityTests.StripInvalidUrlCharsReturnsStrippedString</td><td>Passed</td><td>0h 0m 0s</td></tr>
<tr><td>Tests.AppModelExtensions.UrlUtilityTests.ReplaceInvalidUrlCharsReturnsStrippedString</td><td>Passed</td><td>0h 0m 0s</td></tr>

</table>



