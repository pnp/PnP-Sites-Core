using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("OfficeDevPnP.Core")]
#if SP2013
[assembly: AssemblyDescription("Office Dev PnP Core library for SharePoint 2013")]
#elif SP2016
[assembly: AssemblyDescription("Office Dev PnP Core library for SharePoint 2016")]
#else
[assembly: AssemblyDescription("Office Dev PnP Core library for SharePoint Online")]
#endif
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("OfficeDevPnP.Core")]
[assembly: AssemblyCopyright("Copyright © Microsoft 2018")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]
[assembly: NeutralResourcesLanguage("en-US")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("065331b6-0540-44e1-84d5-d38f09f17f9e")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// Convention:
// Major version = current version 2
// Minor version = Sequence...version 0 was with January release...so 1=Feb 2=Mar...11=Jan 2017...15=May 2017...20=Nov
// Third part = version indenpendant showing the release month in YYMM
// Fourth part = 0 normally or a sequence number when we do an emergency release
[assembly: AssemblyVersion("2.26.1805.1")]
[assembly: AssemblyFileVersion("2.26.1805.1")]

[assembly: InternalsVisibleTo("OfficeDevPnP.Core.Tests")]