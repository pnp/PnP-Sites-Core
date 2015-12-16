using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Operations;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Resources;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions
{
    public class AppManager
    {
        #region Constants
        private static string PubXmlBuildConfiguration = "%BuildConfiguration%";
        private static string PubXmlSiteUrlToLaunchAfterPublish = "%SiteUrlToLaunchAfterPublish%";
        private static string PubXmlServiceURL = "%ServiceURL%";
        private static string PubXmlIisAppPath = "%IisAppPath%";
        private static string PubXmlUserName = "%UserName%";
        private static string PubXmlClientId = "%ClientId%";
        private static string PubXmlClientSecret = "%ClientSecret%";
        private static string PublishingProfileName = "Automation";
        #endregion

        #region Private variables
        private string sharePointUrl;
        private AuthenticationType authenticationType = AuthenticationType.Office365;
        private string userName = "";
        private SecureString password;
        #endregion

        #region Construction
        public AppManager(string sharePointUrl, string credentialManagerLabel)
        {
            this.sharePointUrl = sharePointUrl;
            
            if (sharePointUrl.IndexOf("sharepoint.com") > -1)
            {
                authenticationType = AuthenticationType.Office365;
            }
            else
            {
                authenticationType = AuthenticationType.NetworkCredentials;
            }

            NetworkCredential cred = CredentialManager.GetCredential(credentialManagerLabel);
            userName = cred.UserName;
            password = cred.SecurePassword;
        }

        public AppManager(string sharePointUrl, AuthenticationType authenticationType, string credentialManagerLabel)
        {
            this.sharePointUrl = sharePointUrl;
            this.authenticationType = authenticationType;
            NetworkCredential cred = CredentialManager.GetCredential(credentialManagerLabel);
            userName = cred.UserName;
            password = cred.SecurePassword;        
        }
        public AppManager(string sharePointUrl, string userName, SecureString password)
        {
            this.sharePointUrl = sharePointUrl;

            if (sharePointUrl.IndexOf("sharepoint.com") > -1)
            {
                authenticationType = AuthenticationType.Office365;
            }
            else
            {
                authenticationType = AuthenticationType.NetworkCredentials;
            }

            this.userName = userName;
            this.password = password;
        }

        public AppManager(string sharePointUrl, AuthenticationType authenticationType, string userName, SecureString password)
        {
            this.sharePointUrl = sharePointUrl;
            this.authenticationType = authenticationType;
            this.userName = userName;
            this.password = password;
        }
        #endregion

        #region Properties

        #endregion

        #region Public Methods
        public bool DeployProviderHostedAppAsAzureWebSite(string sharePointProjectFile, string sharePointWebProjectFile, string clientId, string clientSecret, string applicationHost, string siteUrl, string IisAppPath, string azurePublishingSettingsFile, string packageFolder, string buildConfiguration = "Release", string visualStudioVersion="14.0")
        {
            bool createAppPackageResult = false;

            if (String.IsNullOrEmpty(sharePointProjectFile) || !System.IO.File.Exists(sharePointProjectFile))
            {
                throw new ArgumentException(String.Format("Provided SharePoint project file ({0}) is invalid.", sharePointProjectFile));
            }

            if (String.IsNullOrEmpty(sharePointWebProjectFile) || !System.IO.File.Exists(sharePointWebProjectFile))
            {
                throw new ArgumentException(String.Format("Provided SharePoint Web project file ({0}) is invalid.", sharePointWebProjectFile));
            }

            if (String.IsNullOrEmpty(azurePublishingSettingsFile) || !System.IO.File.Exists(azurePublishingSettingsFile))
            {
                throw new ArgumentException(String.Format("Provided Azure publishing settings file ({0}) is invalid.", azurePublishingSettingsFile));
            }
            

            if (String.IsNullOrEmpty(clientId))
            {
                throw new ArgumentException("Please provide a client id");
            }

            if (String.IsNullOrEmpty(clientSecret))
            {
                throw new ArgumentException("Please provide a client secret");
            }

            if (String.IsNullOrEmpty(applicationHost))
            {
                throw new ArgumentException("Please provide an application host");
            }

            if (String.IsNullOrEmpty(siteUrl))
            {
                throw new ArgumentException("Please provide an site Url");
            }

            if (String.IsNullOrEmpty(buildConfiguration))
            {
                throw new ArgumentException("Please provide a build configuration (e.g. release or debug)");
            }

            if (String.IsNullOrEmpty(visualStudioVersion))
            {
                throw new ArgumentException("Please provide a Visual Studio version (e.g. 12.0 or 14.0)");
            }

            // update web.config - clientid and secret

            // read azure publishing file, grab username and password


            // Get a base template that will be used for the publishing
            string publishingTemplateString = ResourceManager.GetPublishingXmlTemplate(true, true, PublishingTypes.AzureWebSite);

            // replace tokens in the publishing template
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlClientId, clientId);
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlClientSecret, clientSecret);
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlBuildConfiguration, buildConfiguration);
            // Applicationhost is like bjansen-automation1.scm.azurewebsites.net:443                        
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlServiceURL, applicationHost);
            // Site Url to launch after publish is like https://bjansen-automation1.azurewebsites.net
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlSiteUrlToLaunchAfterPublish, siteUrl);
            // IIS app path is like bjansen-automation1
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlIisAppPath, IisAppPath);

            // insert the publishing template in the solution
            string publishingXmlFile = "";
            try
            {
                // We need a new publishing xml file name
                publishingXmlFile = String.Format("{0}\\Properties\\PublishProfiles\\{1}.pubxml", Path.GetDirectoryName(sharePointWebProjectFile), AppManager.PublishingProfileName);

                // Add a new publishing xml file
                if (!Directory.Exists(Path.GetDirectoryName(publishingXmlFile)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(publishingXmlFile));
                }
                File.WriteAllText(publishingXmlFile, publishingTemplateString);

                // Trigger package build
                Hashtable packageBuildParameters = new Hashtable();
                packageBuildParameters.Add("ProjectFile", sharePointProjectFile);

                // Package folder need to end with a \
                if (!packageFolder.EndsWith(@"\"))
                {
                    packageFolder = packageFolder + @"\";
                }

                if (Directory.Exists(packageFolder))
                {
                    Console.WriteLine("Package directory already exists...it will be removed as part of the build process");
                }

                packageBuildParameters.Add("OutputPath", packageFolder);
                // we pass the visual studio version: important that to know that the visual studio version defined in the
                // project file is also installed. Ideally the passed version and the one define in the project match
                packageBuildParameters.Add("VisualStudioVersion", visualStudioVersion);
                packageBuildParameters.Add("BuildConfiguration", buildConfiguration);
                // override the publishing profile that's needed for packaging the app
                packageBuildParameters.Add("ActivePublishProfile", AppManager.PublishingProfileName);

                string packageBuildResult = Run.RunScript(String.Format(@"{0}\Scripts\PublishProviderHosted.ps1", ResourceManager.GetAssemblyDirectory()), packageBuildParameters);

                if (packageBuildResult.Contains("Build FAILED"))
                {
                    Console.WriteLine("App Package creation failed");
                    Console.WriteLine(packageBuildResult);
                }
                else
                {
                    createAppPackageResult = true;
                }
            }
            finally
            {
                // delete the created publishing xml file to restore the project back in it's original state
                File.Delete(publishingXmlFile);
            }
            
            return createAppPackageResult;
        }

        public bool CreateAppPackageForProviderHostedApp(string sharePointProjectFile, string sharePointWebProjectFile, string clientId, string siteUrl, string packageFolder, out string appPackageName, string buildConfiguration = "Release", string visualStudioVersion = "14.0")
        {
            bool createAppPackageResult = false;

            if (String.IsNullOrEmpty(sharePointProjectFile) || !System.IO.File.Exists(sharePointProjectFile))
            {
                throw new ArgumentException(String.Format("Provide SharePoint project file ({0}) is invalid.", sharePointProjectFile));
            }

            if (String.IsNullOrEmpty(sharePointWebProjectFile) || !System.IO.File.Exists(sharePointWebProjectFile))
            {
                throw new ArgumentException(String.Format("Provide SharePoint Web project file ({0}) is invalid.", sharePointWebProjectFile));
            }

            if (String.IsNullOrEmpty(clientId))
            {
                throw new ArgumentException("Please provide a client id");
            }

            if (String.IsNullOrEmpty(siteUrl))
            {
                throw new ArgumentException("Please provide an site Url");
            }

            if (String.IsNullOrEmpty(buildConfiguration))
            {
                throw new ArgumentException("Please provide a build configuration (e.g. release or debug)");
            }

            if (String.IsNullOrEmpty(visualStudioVersion))
            {
                throw new ArgumentException("Please provide a Visual Studio version (e.g. 12.0 or 14.0)");
            }

            // Get a base template that will be used for the publishing - just take Azure template...doesn't matter since 
            // we're only creating the app package
            string publishingTemplateString = ResourceManager.GetPublishingXmlTemplate(true, true, PublishingTypes.AzureWebSite);

            // replace tokens in the publishing template
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlClientId, clientId);
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlBuildConfiguration, buildConfiguration);
            // Site Url to launch after publish is like https://bjansen-automation1.azurewebsites.net
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlSiteUrlToLaunchAfterPublish, siteUrl);

            // insert the publishing template in the solution
            string publishingXmlFile = "";
            try
            {
                // We need a new publishing xml file name
                publishingXmlFile = String.Format("{0}\\Properties\\PublishProfiles\\{1}.pubxml", Path.GetDirectoryName(sharePointWebProjectFile), AppManager.PublishingProfileName);

                // Add a new publishing xml file
                if (!Directory.Exists(Path.GetDirectoryName(publishingXmlFile)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(publishingXmlFile));
                }
                File.WriteAllText(publishingXmlFile, publishingTemplateString);

                // Trigger package build
                Hashtable packageBuildParameters = new Hashtable();
                packageBuildParameters.Add("ProjectFile", sharePointProjectFile);

                // Package folder need to end with a \
                if (!packageFolder.EndsWith(@"\"))
                {
                    packageFolder = packageFolder + @"\";
                }

                if (Directory.Exists(packageFolder))
                {
                    Console.WriteLine("Package directory already exists...it will be removed as part of the build process");
                }

                packageBuildParameters.Add("OutputPath", packageFolder);
                // we pass the visual studio version: important that to know that the visual studio version defined in the
                // project file is also installed. Ideally the passed version and the one define in the project match
                packageBuildParameters.Add("VisualStudioVersion", visualStudioVersion);
                packageBuildParameters.Add("BuildConfiguration", buildConfiguration);
                // override the publishing profile that's needed for packaging the app
                packageBuildParameters.Add("ActivePublishProfile", AppManager.PublishingProfileName);

                string packageBuildResult = Run.RunScript(String.Format(@"{0}\Scripts\GenerateAppPackageProviderHosted.ps1", ResourceManager.GetAssemblyDirectory()), packageBuildParameters);

                if (packageBuildResult.Contains("Build FAILED"))
                {
                    Console.WriteLine("App Package creation failed");
                    Console.WriteLine(packageBuildResult);
                    appPackageName = "";
                }
                else
                {
                    createAppPackageResult = true;
                    appPackageName = GetAppPackageFile(packageFolder);
                    // Remove all unneeded files from the package folder
                    CleanupAppPackageFolder(packageFolder);
                }
            }
            finally
            {
                // delete the created publishing xml file to restore the project back in it's original state
                File.Delete(publishingXmlFile);
            }

            return createAppPackageResult;
        }

        public bool CreateAppPackageForSharePointHostedApp(string sharePointProjectFile, string packageFolder, out string appPackageName, string buildConfiguration = "Release", string visualStudioVersion = "14.0")
        {
            bool createAppPackageResult = false;

            if (String.IsNullOrEmpty(sharePointProjectFile) || !System.IO.File.Exists(sharePointProjectFile))
            {
                throw new ArgumentException(String.Format("Provide SharePoint project file ({0}) is invalid.", sharePointProjectFile));
            }

            if (String.IsNullOrEmpty(buildConfiguration))
            {
                throw new ArgumentException("Please provide a build configuration (e.g. release or debug)");
            }

            if (String.IsNullOrEmpty(visualStudioVersion))
            {
                throw new ArgumentException("Please provide a Visual Studio version (e.g. 12.0 or 14.0)");
            }

            try
            {
                // Trigger package build
                Hashtable packageBuildParameters = new Hashtable();
                packageBuildParameters.Add("ProjectFile", sharePointProjectFile);

                // Package folder need to end with a \
                if (!packageFolder.EndsWith(@"\"))
                {
                    packageFolder = packageFolder + @"\";
                }

                if (Directory.Exists(packageFolder))
                {
                    Console.WriteLine("Package directory already exists...it will be removed as part of the build process");
                }

                packageBuildParameters.Add("OutputPath", packageFolder);
                // we pass the visual studio version: important that to know that the visual studio version defined in the
                // project file is also installed. Ideally the passed version and the one define in the project match
                packageBuildParameters.Add("VisualStudioVersion", visualStudioVersion);
                packageBuildParameters.Add("BuildConfiguration", buildConfiguration);

                string packageBuildResult = Run.RunScript(String.Format(@"{0}\Scripts\GenerateAppPackageSharePointHosted.ps1", ResourceManager.GetAssemblyDirectory()), packageBuildParameters);

                if (packageBuildResult.Contains("Build FAILED"))
                {
                    Console.WriteLine("App Package creation failed");
                    Console.WriteLine(packageBuildResult);
                    appPackageName = "";
                }
                else
                {
                    createAppPackageResult = true;
                    appPackageName = GetAppPackageFile(packageFolder);
                    // Remove all unneeded files from the package folder
                    CleanupAppPackageFolder(packageFolder);
                }
            }
            finally
            {
            }

            return createAppPackageResult;
        }

        public bool RegisterApplication(ref string clientId, ref string clientSecret, string title, string appDomain, string redirectUri)
        {
            bool appRegNewResult = false;

            if (String.IsNullOrEmpty(sharePointUrl) || !ValidateUri(sharePointUrl))
            {
                throw new ArgumentException("Please provide a valid value for the SharePoint url");
            }

            if (String.IsNullOrEmpty(title))
            {
                throw new ArgumentException("Please provide a value for title");
            }

            if (String.IsNullOrEmpty(appDomain))
            {
                throw new ArgumentException("Please provide a value for appDomain");
            }

            if (String.IsNullOrEmpty(redirectUri) || !ValidateUri(redirectUri))
            {
                throw new ArgumentException("Please provide a valid value for the redirect uri");
            }

            if (clientId == null)
            {
                clientId = AppRegNew.GenerateAppId();
            }


            if (String.IsNullOrEmpty(clientSecret))
            {
                clientSecret = AppRegNew.GenerateAppSecret();
            }

            AppRegNew appRegNew = new AppRegNew(sharePointUrl, this.authenticationType, userName, password);
            appRegNew.AppId = clientId;
            appRegNew.AppSecret = clientSecret;
            appRegNew.Title = title;
            appRegNew.HostUri = appDomain;
            appRegNew.RedirectUri = redirectUri;
            appRegNewResult = appRegNew.Execute(this.authenticationType == AuthenticationType.Office365 ? "online" : "onpremises");

            Thread.Sleep(1000);
            if (appRegNewResult == true)
            {
                Console.WriteLine("App Registration done for \n \t App Id: " + appRegNew.AppId
                     + "\n \t App Secret: " + appRegNew.AppSecret
                     + "\n \t App Title: " + appRegNew.Title
                     + "\n \t Host Uri: " + appRegNew.HostUri
                     + "\n \t Redirect Uri: " + appRegNew.RedirectUri);
                Console.WriteLine("App Registered Successfully");
            }

            else
            {
                Console.WriteLine("App Registration failed for below app details \n \t App Id: " + appRegNew.AppId
                     + "\n \t App Secret: " + appRegNew.AppSecret
                     + "\n \t App Title: " + appRegNew.Title
                     + "\n \t Host Uri: " + appRegNew.HostUri
                     + "\n \t Redirect Uri: " + appRegNew.RedirectUri);
            }
            return appRegNewResult;
        }
        #endregion

        #region Private methods

        private void CleanupAppPackageFolder(string packageFolder)
        {
            string[] files = Directory.GetFiles(packageFolder);

            // clean up in root folder
            for (int i = 0; i < files.Length; i++)
            {
                File.Delete(files[i]);
            }

            // cleanup directories
            string[] directories = Directory.GetDirectories(packageFolder);
            for (int i = 0; i < directories.Length; i++)
            {
                if (!directories[i].Equals(Path.Combine(packageFolder, "app.publish"), StringComparison.InvariantCultureIgnoreCase))
                {
                    Directory.Delete(directories[i], true);
                }
            }

            // cleanup in app.publish folder
            files = Directory.GetFiles(Path.Combine(packageFolder, "app.publish"), "*.*", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                if (!files[i].EndsWith(".app"))
                {
                    if (File.Exists(files[i]))
                    {
                        File.Delete(files[i]);
                    }
                }
            }
        }

        private string GetAppPackageFile(string packageFolder)
        {
            string appPackageFile = null;

            string[] appPackages = Directory.GetFiles(Path.Combine(packageFolder, "app.publish"), "*.app", SearchOption.AllDirectories);

            if (appPackages.Length > 0)
            {
                appPackageFile = appPackages[0];
            }

            return appPackageFile;
        }


        private bool ValidateUri(string url)
        {            
            url = url.Trim();

            if (url.Length == 0)
            {
                return false;
            }

            Uri uri;
            if (!Uri.TryCreate(url, UriKind.Absolute, out uri))
            {
                return false;
            }

            if (uri.Scheme != Uri.UriSchemeHttp &&
                uri.Scheme != Uri.UriSchemeHttps)
            {
                return false;
            }

            return true;
        }

        #region Not needed anymore, but saving for a while
        //private string SetActivePublishingProfile(XmlDocument xDoc, string sharePointProjectFile, string publishingProfile)
        //{
        //    string previousValue = null;
            
        //    //We've a namespace to take in account
        //    XmlNamespaceManager nsmgr = new XmlNamespaceManager(xDoc.NameTable);
        //    nsmgr.AddNamespace("ns", xDoc.DocumentElement.NamespaceURI);

        //    // Query the ActivePublishProfile property node
        //    XmlNode activePublishProfile = xDoc.DocumentElement.SelectSingleNode("/ns:Project/ns:PropertyGroup[1]/ns:ActivePublishProfile", nsmgr);

        //    if (activePublishProfile != null && !String.IsNullOrEmpty(activePublishProfile.InnerText))
        //    {
        //        // we've a publishing profile value
        //        if (publishingProfile != null)
        //        {
        //            // The node was there, so let's update

        //            // store the previous value as we need to restore it afterwards
        //            previousValue = activePublishProfile.InnerText;
        //            // update the node with the new value
        //            activePublishProfile.InnerText = publishingProfile;
        //        }
        //        else
        //        {
        //            // there's no publishing profile value...happens when there was none set and we're now restoring the old settings --> we need to remove the node
        //            activePublishProfile.ParentNode.RemoveChild(activePublishProfile);
        //        }
        //    }
        //    else
        //    {
        //        // The node was not there, let's add it

        //        if (publishingProfile != null)
        //        {
        //            // Query the ActivePublishProfile property node
        //            XmlNode propertyGroup = xDoc.DocumentElement.SelectSingleNode("/ns:Project/ns:PropertyGroup[1]", nsmgr);
        //            XmlElement activePublishingProfileNode = xDoc.CreateElement("ActivePublishProfile", xDoc.DocumentElement.NamespaceURI);
        //            activePublishingProfileNode.InnerText = publishingProfile;
        //            propertyGroup.AppendChild(activePublishingProfileNode);
        //        }
        //        else
        //        {
        //            throw new ArgumentException("Provided publishing profile is null which is not a supported case.");
        //        }
        //    }

        //    // persist the changes
        //    xDoc.Save(sharePointProjectFile);

        //    return previousValue;
        //}

        //private string GetAvailablePublishingXml(string sharePointProjectFile, string sharePointWebProjectFile)
        //{
        //    string publishingXmlFile = "";

        //    if (String.IsNullOrEmpty(sharePointProjectFile) || !System.IO.File.Exists(sharePointProjectFile))
        //    {
        //        throw new ArgumentException(String.Format("Provide SharePoint project file ({0}) is invalid.", sharePointProjectFile));
        //    }

        //    if (String.IsNullOrEmpty(sharePointWebProjectFile) || !System.IO.File.Exists(sharePointWebProjectFile))
        //    {
        //        throw new ArgumentException(String.Format("Provide SharePoint Web project file ({0}) is invalid.", sharePointWebProjectFile));
        //    }

        //    // e.g. C:\GitHub\BertPnP\Samples\Core.EmbedJavaScript\Core.EmbedJavaScriptWeb\Properties\PublishProfiles\bert2.pubxml
        //    string folder = String.Format("{0}\\Properties\\PublishProfiles", Path.GetDirectoryName(sharePointWebProjectFile));

        //    // Load the project file as we need to understand the 
        //    //XmlElement sharePointProjectFileXml = LoadXmlFile(sharePointProjectFile);


        //    // Get the pubxml files in this folder
        //    string[] publishingXmlFiles = Directory.GetFiles(folder, "*.pubxml");
        //    if (publishingXmlFiles.Length == 1)
        //    {
        //        publishingXmlFile = publishingXmlFiles[0];
        //    }
        //    else if (publishingXmlFiles.Length > 1)
        //    {
        //        //There are multiple publish xml files in the project, figure out which one is the active one by looking at the SharePoint project properties
        //        XmlElement sharePointProjectFileXml = LoadXmlFile(sharePointProjectFile);

        //        //We've a namespace to take in account
        //        XmlNamespaceManager nsmgr = new XmlNamespaceManager(sharePointProjectFileXml.OwnerDocument.NameTable);
        //        nsmgr.AddNamespace("ns", sharePointProjectFileXml.OwnerDocument.DocumentElement.NamespaceURI);

        //        // Query the ActivePublishProfile property node
        //        XmlNode activePublishProfile = sharePointProjectFileXml.SelectSingleNode("/ns:Project/ns:PropertyGroup[1]/ns:ActivePublishProfile", nsmgr);

        //        if (activePublishProfile != null && !String.IsNullOrEmpty(activePublishProfile.InnerText))
        //        {
        //            publishingXmlFile = String.Format("{0}\\{1}.pubxml", folder, activePublishProfile.InnerText);
        //        }
        //    }

        //    return publishingXmlFile;
        //}

        //private XmlElement LoadXmlFile(string fileName)
        //{
        //    if (File.Exists(fileName))
        //    {
        //        XmlDocument xDoc = new XmlDocument();
        //        xDoc.Load(fileName);

        //        // If there's a namespace, then add it 
        //        return xDoc.DocumentElement;
        //    }
        //    else
        //    {
        //        throw new FileNotFoundException(String.Format("XML file {0} was not found", fileName));
        //    }
        //}
        #endregion
        #endregion

    }
}
