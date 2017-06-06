# PnP Build and Test automation #

### Summary ###
This project contains the build scripts and build and test extensions used to automate the building and testing of PnP

### Applies to ###
-  Office 365 Multi Tenant (MT)
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises

### Prerequisites ###
Git needs to be installed, see documentation for details

### Solution ###
Solution | Author(s)
---------|----------
OfficeDevPnP.Core.Tools.UnitTest | Bert Jansen (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | July 4th 2015 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Purpose #
The purpose of this project is to automate testing of PnP against multiple environments using multiple configurations. In the current setup this tool is used to execute the PnP unit tests against various SharePoint Online (Office 365 MT) and SharePoint 2013/2016 on-premises environments. For each environment the unit test execution in running for 2 configurations: username + password and app-only.

![](http://i.imgur.com/HPUvUJg.png)

The test engine is a console application that uses an regular MSBuild script in combination with a set of custom MSBuild tasks and a custom VS Test logger to automate the test execution.

# How to use #
Below are the steps needed to get this solution working.

## Prepare a SQL Azure DB ##
The PnP test automation tool logs it's result into a SQL Azure database. This same database is also used to define the test configurations to run. You can find the schema of this database in the `OfficeDevPnP.Core.Tools.UnitTest.SQL` project.

## Copy needed files to the build server ##
Following files are needed:
- Copy the output from the release build of project `OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner` to a location (e.g. folder `c:\pnpunittestrunner`) on the server that's running the build automation. 
- Copy the PnPSQLCore.targets file 

## Create a .bat file that runs OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner.exe ##
A possible bat file can be the following (is also copied as part of the build output)

```batch
\\dc1\Admin\pnptestrunner\OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner.exe "AzureVMSP2013Credentials" "data source=tcp:yourdb.database.windows.net,1433;Database=PnP;User ID=pnp;Password=yourpwd;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;MultipleActiveResultSets=True;App=EntityFramework" "\\dc1\Admin\pnptestrunner\PnPSQLCore.targets"
timeout 30
\\dc1\Admin\pnptestrunner\OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner.exe "AzureVMSP2013AppOnly" "data source=tcp:yourdb.database.windows.net,1433;Database=PnP;User ID=pnp;Password=yourpwd;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;MultipleActiveResultSets=True;App=EntityFramework" "\\dc1\Admin\pnptestrunner\PnPSQLCore.targets"
timeout 30
\\dc1\Admin\pnptestrunner\OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner.exe "MTFirstReleaseCredentials" "data source=tcp:yourdb.database.windows.net,1433;Database=PnP;User ID=pnp;Password=yourpwd;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;MultipleActiveResultSets=True;App=EntityFramework" "\\dc1\Admin\pnptestrunner\PnPSQLCore.targets"
timeout 30
\\dc1\Admin\pnptestrunner\OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner.exe "MTFirstReleaseAppOnly" "data source=tcp:yourdb.database.windows.net,1433;Database=PnP;User ID=pnp;Password=yourpwd;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;MultipleActiveResultSets=True;App=EntityFramework" "\\dc1\Admin\pnptestrunner\PnPSQLCore.targets"
timeout 30
```

Note that in this bat file we provide input parameters to the `PnPTestRunner.exe` console application:
- **Configuration to run**: this is the test configuration defined in the Azure DB that needs to be executed
- **Connection string** to the SQL Azure database
- **MSBuild file** (.targets file) to execute

## Copy the OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.dll to the VS extensions folder ##
We're using a custom VS test log writer that writes output to MD. This log writer needs to be copied to `C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\CommonExtensions\Microsoft\TestWindow\Extensions`. Replace the Visual Studio folder with the version you're using.

## Update the PnPSQLCore.Targets file to suit your needs ##
The below sections need to be adjusted to match your environment:

```XML
<!-- PnP Repo information -->
<PropertyGroup Label="PnP">
  <PnPRepo>c:\temp\pnpbuild</PnPRepo>
  <PnPRepoUrl>https://github.com/OfficeDev/PnP-Sites-Core.git</PnPRepoUrl>
</PropertyGroup>

<!-- Unit test information-->
<PropertyGroup Label="Test information">
  <ConfigurationPath>C:\GitHub\BertPnPSitesCore\Core\Tools\OfficeDevPnP.Core.Tools.UnitTest\OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions</ConfigurationPath>
  <PnPExtensionsAssembly>$(ConfigurationPath)\bin\Debug\OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.dll</PnPExtensionsAssembly>
  <ConfigurationFile>mastertestconfiguration.xml</ConfigurationFile>
  <TestResultsPath>$(PnPRepo)temp</TestResultsPath>
  <VSTestExe>C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe</VSTestExe>
  <VSTestExtensionPath>C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\CommonExtensions\Microsoft\TestWindow\Extensions</VSTestExtensionPath>
</PropertyGroup>
```

These are the important parameters to change:
- **PnPRepo**: this defines where the PnP repo will be pulled down. The build scripts will use a separate copy, **so please do not put your working PnP fork/clone here**
- **ConfigurationPath**: this is the folder in which you've copied all the files needed for the test automation

## Setup git ##
The build script requires git to be present, hence git needs to be installed. The tested version is git for windows which can be fetched from here: http://msysgit.github.io/. If you want to push back changes to the PnP repo than ensure git is properly configured.

## Create a scheduled task ##
Final step is creating a scheduled task that executes the created bat file on a regular basis.

<img src="https://telemetry.sharepointpnp.com/pnp-sites-core/core/tools/OfficeDevPnP.Core.Tools.UnitTest" /> 
