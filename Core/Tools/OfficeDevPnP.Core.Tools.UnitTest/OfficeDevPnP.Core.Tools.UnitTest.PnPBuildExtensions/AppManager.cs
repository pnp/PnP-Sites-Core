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
        public bool CreateAppPackageForProviderHostedApp(string sharePointProjectFile, string sharePointWebProjectFile, string clientId, string clientSecret, string applicationHost, string packageFolder, out string appPackageName, string buildConfiguration = "Release")
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

            if (String.IsNullOrEmpty(clientSecret))
            {
                throw new ArgumentException("Please provide a client secret");
            }

            if (String.IsNullOrEmpty(applicationHost))
            {
                throw new ArgumentException("Please provide an application host");
            }

            // Get a base template that will be used for the publishing
            string publishingTemplateString = ResourceManager.GetPublishingXmlTemplate(true, true, PublishingTypes.AzureWebSite);

            // replace tokens in the publishing template
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlClientId, clientId);
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlClientSecret, clientSecret);
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlServiceURL, applicationHost);
            publishingTemplateString = publishingTemplateString.Replace(AppManager.PubXmlBuildConfiguration, buildConfiguration);

            // insert the publishing template in the solution
            //bool publishingXmlWasDefined = false;
            string publishingXmlFile = "";
            string currentActivePublishingProfile = null;
            XmlDocument xDoc = null;
            try
            {
                //// Do we already have a publishing XML defined?
                //publishingXmlFile = GetAvailablePublishingXml(sharePointProjectFile, sharePointWebProjectFile);

                //// We've a publishing xml file in the project, let's update it
                //if (!String.IsNullOrEmpty(publishingXmlFile))
                //{
                //    publishingXmlWasDefined = true;

                //    // Rename the original active publishing XML file
                //    File.Move(publishingXmlFile, String.Format("{0}.old", publishingXmlFile));
                //}
                //else
                //{
                //    // We need a new publishing xml file name
                //    publishingXmlFile = String.Format("{0}\\Properties\\PublishProfiles\\automation.pubxml", Path.GetDirectoryName(sharePointWebProjectFile));
                //}


                // We need a new publishing xml file name
                publishingXmlFile = String.Format("{0}\\Properties\\PublishProfiles\\{1}.pubxml", Path.GetDirectoryName(sharePointWebProjectFile), AppManager.PublishingProfileName);

                // Load SharePoint Project file to find/set the ActivePublishingProfile
                xDoc = new XmlDocument();
                xDoc.Load(sharePointProjectFile);
                currentActivePublishingProfile = SetActivePublishingProfile(xDoc, sharePointProjectFile, AppManager.PublishingProfileName);

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
                string packageBuildResult = Run.RunScript(String.Format(@"{0}\Scripts\GenerateAppPackage.ps1", ResourceManager.GetAssemblyDirectory()), packageBuildParameters);

                if (!packageBuildResult.Contains("Build FAILED"))
                {
                    createAppPackageResult = true;
                    appPackageName = "todo";
                }
                else
                {
                    Console.WriteLine("App Package creation failed");
                    Console.WriteLine(packageBuildResult);
                    appPackageName = "";
                }
            }
            finally
            {
                // delete the created publishing xml file to restore the project back in it's original state
                File.Delete(publishingXmlFile);

                // restore the project file to it's original state. 
                SetActivePublishingProfile(xDoc, sharePointProjectFile, currentActivePublishingProfile);

                //if (publishingXmlWasDefined)
                //{
                //    // revert back to the original publishing XML if there was one defined
                //    File.Move(String.Format("{0}.old", publishingXmlFile), publishingXmlFile);
                //}
                
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

        private string SetActivePublishingProfile(XmlDocument xDoc, string sharePointProjectFile, string publishingProfile)
        {
            string previousValue = null;
            
            //We've a namespace to take in account
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xDoc.NameTable);
            nsmgr.AddNamespace("ns", xDoc.DocumentElement.NamespaceURI);

            // Query the ActivePublishProfile property node
            XmlNode activePublishProfile = xDoc.DocumentElement.SelectSingleNode("/ns:Project/ns:PropertyGroup[1]/ns:ActivePublishProfile", nsmgr);

            if (activePublishProfile != null && !String.IsNullOrEmpty(activePublishProfile.InnerText))
            {
                // we've a publishing profile value
                if (publishingProfile != null)
                {
                    // The node was there, so let's update

                    // store the previous value as we need to restore it afterwards
                    previousValue = activePublishProfile.InnerText;
                    // update the node with the new value
                    activePublishProfile.InnerText = publishingProfile;
                }
                else
                {
                    // there's no publishing profile value...happens when there was none set and we're now restoring the old settings --> we need to remove the node
                    activePublishProfile.ParentNode.RemoveChild(activePublishProfile);
                }
            }
            else
            {
                // The node was not there, let's add it

                if (publishingProfile != null)
                {
                    // Query the ActivePublishProfile property node
                    XmlNode propertyGroup = xDoc.DocumentElement.SelectSingleNode("/ns:Project/ns:PropertyGroup[1]", nsmgr);
                    XmlElement activePublishingProfileNode = xDoc.CreateElement("ActivePublishingProfile", xDoc.DocumentElement.NamespaceURI);
                    activePublishingProfileNode.InnerText = publishingProfile;
                    propertyGroup.AppendChild(activePublishingProfileNode);
                }
                else
                {
                    throw new ArgumentException("Provided publishing profile is null which is not a supported case.");
                }
            }

            // persist the changes
            xDoc.Save(sharePointProjectFile);

            return previousValue;
        }

        private string GetAvailablePublishingXml(string sharePointProjectFile, string sharePointWebProjectFile)
        {
            string publishingXmlFile = "";

            if (String.IsNullOrEmpty(sharePointProjectFile) || !System.IO.File.Exists(sharePointProjectFile))
            {
                throw new ArgumentException(String.Format("Provide SharePoint project file ({0}) is invalid.", sharePointProjectFile));
            }

            if (String.IsNullOrEmpty(sharePointWebProjectFile) || !System.IO.File.Exists(sharePointWebProjectFile))
            {
                throw new ArgumentException(String.Format("Provide SharePoint Web project file ({0}) is invalid.", sharePointWebProjectFile));
            }

            // e.g. C:\GitHub\BertPnP\Samples\Core.EmbedJavaScript\Core.EmbedJavaScriptWeb\Properties\PublishProfiles\bert2.pubxml
            string folder = String.Format("{0}\\Properties\\PublishProfiles", Path.GetDirectoryName(sharePointWebProjectFile));

            // Load the project file as we need to understand the 
            //XmlElement sharePointProjectFileXml = LoadXmlFile(sharePointProjectFile);


            // Get the pubxml files in this folder
            string[] publishingXmlFiles = Directory.GetFiles(folder, "*.pubxml");
            if (publishingXmlFiles.Length == 1)
            {
                publishingXmlFile = publishingXmlFiles[0];
            }
            else if (publishingXmlFiles.Length > 1)
            {
                //There are multiple publish xml files in the project, figure out which one is the active one by looking at the SharePoint project properties
                XmlElement sharePointProjectFileXml = LoadXmlFile(sharePointProjectFile);

                //We've a namespace to take in account
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(sharePointProjectFileXml.OwnerDocument.NameTable);
                nsmgr.AddNamespace("ns", sharePointProjectFileXml.OwnerDocument.DocumentElement.NamespaceURI);

                // Query the ActivePublishProfile property node
                XmlNode activePublishProfile = sharePointProjectFileXml.SelectSingleNode("/ns:Project/ns:PropertyGroup[1]/ns:ActivePublishProfile", nsmgr);

                if (activePublishProfile != null && !String.IsNullOrEmpty(activePublishProfile.InnerText))
                {
                    publishingXmlFile = String.Format("{0}\\{1}.pubxml", folder, activePublishProfile.InnerText);
                }
            }

            return publishingXmlFile;
        }

        private XmlElement LoadXmlFile(string fileName)
        {
            if (File.Exists(fileName))
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(fileName);

                // If there's a namespace, then add it 
                return xDoc.DocumentElement;
            }
            else
            {
                throw new FileNotFoundException(String.Format("XML file {0} was not found", fileName));
            }
        }
        #endregion

    }
}
