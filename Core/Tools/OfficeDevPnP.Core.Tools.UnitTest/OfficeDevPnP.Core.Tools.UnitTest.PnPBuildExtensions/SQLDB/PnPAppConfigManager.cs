using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Utilities;
using OfficeDevPnP.Core.Tools.UnitTest.SQL;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQLDB
{
    public class PnPAppConfigManager
    {
        private int configurationId;
        private string sqlConnectionString;

        public PnPAppConfigManager(string sqlConnectionString, string configurationName)
        {
            //System.Diagnostics.Debugger.Launch();
            
            // let's find the right id based on the name
            using (TestModelContainer context = new TestModelContainer(sqlConnectionString))
            {
                var configuration = context.TestConfigurationSet.Where(s => s.Name.Equals(configurationName, StringComparison.InvariantCultureIgnoreCase)).First();
                if (configuration != null)
                {
                    this.configurationId = configuration.Id;
                }
                else
                {
                    throw new Exception(String.Format("Test configuration with name {0} was not found", configurationName));
                }
            }
            this.sqlConnectionString = sqlConnectionString;
        }

        public string GetConfigurationElement(string element)
        {
            using (TestModelContainer context = new TestModelContainer(sqlConnectionString))
            {
                TestConfiguration testConfig = context.TestConfigurationSet.Find(configurationId);
                if (testConfig == null)
                {
                    throw new Exception(String.Format("Test configuration with ID {0} was not found", configurationId));
                }

                if (element.Equals("PnPBranch", StringComparison.InvariantCultureIgnoreCase))
                {
                    return testConfig.Branch;
                }
                else if (element.Equals("PnPBuild", StringComparison.InvariantCultureIgnoreCase))
                {
                    return testConfig.VSBuildConfiguration;
                }

                return "";
            }
        }

        public void GenerateAppConfig(string appConfigFolder)
        {
            using (TestModelContainer context = new TestModelContainer(sqlConnectionString))
            {
                TestConfiguration testConfig = context.TestConfigurationSet.Find(configurationId);
                if (testConfig == null)
                {
                    throw new Exception(String.Format("Test configuration with ID {0} was not found", configurationId));
                }

                string appConfigFile = Path.Combine(appConfigFolder, "app.config");

                // If there's already an app.config file then delete it
                if (File.Exists(appConfigFile))
                {
                    File.Delete(appConfigFile);
                }

                // Generate app.config XML file
                using (XmlWriter writer = XmlWriter.Create(appConfigFile))
                {
                    writer.WriteStartElement("configuration");
                    writer.WriteStartElement("appSettings");

                    // These app settings property value pairs are always present
                    WriteProperty(writer, "SPOTenantUrl", testConfig.TenantUrl);
                    WriteProperty(writer, "SPODevSiteUrl", testConfig.TestSiteUrl);

                    if (testConfig.Type == TestConfigurationType.SharePoint2013 || testConfig.Type == TestConfigurationType.SharePoint2016)
                    {
                        WriteProperty(writer, "SPOCredentialManagerLabel", testConfig.TestAuthentication.CredentialManagerLabel);
                        if (!testConfig.TestAuthentication.AppOnly)
                        {
                            if (!String.IsNullOrEmpty(testConfig.TestAuthentication.CredentialManagerLabel))
                            {
                                NetworkCredential cred = CredentialManager.GetCredential(testConfig.TestAuthentication.CredentialManagerLabel);
                                if (cred.UserName.IndexOf("\\") > 0)
                                {
                                    string[] userParts = cred.UserName.Split('\\');
                                    WriteProperty(writer, "OnPremUserName", userParts[1]);
                                    WriteProperty(writer, "OnPremDomain", userParts[0]);
                                }
                                else
                                {
                                    throw new ArgumentException(String.Format("Username {0} stored in credential manager value {1} needs to be formatted as domain\\user", cred.UserName, testConfig.TestAuthentication.CredentialManagerLabel));
                                }
                            }
                            else
                            {
                                WriteProperty(writer, "OnPremUserName", testConfig.TestAuthentication.User);
                                WriteProperty(writer, "OnPremDomain", testConfig.TestAuthentication.Domain);
                                WriteProperty(writer, "OnPremPassword", testConfig.TestAuthentication.Password);
                            }
                        }
                        else // App-Only
                        {
                            WriteProperty(writer, "AppId", testConfig.TestAuthentication.AppId);
                            WriteProperty(writer, "AppSecret", testConfig.TestAuthentication.AppSecret);
                        }

                        // dump additional properties
                        foreach (var testConfigurationProperty in testConfig.TestConfigurationProperties)
                        {
                            WriteProperty(writer, testConfigurationProperty.Name, testConfigurationProperty.Value);
                        }

                        // dump "special" additional properties
                        WriteProperty(writer, "TestAutomationDatabaseConnectionString", GetConnectionString(sqlConnectionString));
                    }
                    else // Online
                    {
                        WriteProperty(writer, "SPOCredentialManagerLabel", testConfig.TestAuthentication.CredentialManagerLabel);
                        if(!testConfig.TestAuthentication.AppOnly)
                        {
                            // System.Diagnostics.Debugger.Launch();
                            // Always output the username since some tests depend on this
                            if (!String.IsNullOrEmpty(testConfig.TestAuthentication.CredentialManagerLabel))
                            {
                                NetworkCredential cred = CredentialManager.GetCredential(testConfig.TestAuthentication.CredentialManagerLabel);
                                WriteProperty(writer, "SPOUserName", cred.UserName);
                            }
                            else
                            {
                                WriteProperty(writer, "SPOUserName", testConfig.TestAuthentication.User);
                                WriteProperty(writer, "SPOPassword", testConfig.TestAuthentication.Password);
                            }
                        }
                        else // App-Only
                        {
                            WriteProperty(writer, "AppId", testConfig.TestAuthentication.AppId);
                            WriteProperty(writer, "AppSecret", testConfig.TestAuthentication.AppSecret);
                        }
                        
                        // dump additional properties
                        foreach (var testConfigurationProperty in testConfig.TestConfigurationProperties)
                        {
                            WriteProperty(writer, testConfigurationProperty.Name, testConfigurationProperty.Value);
                        }

                        // dump "special" additional properties
                        WriteProperty(writer, "TestAutomationDatabaseConnectionString", GetConnectionString(sqlConnectionString));

                    }
                    writer.WriteEndElement(); //appSettings

                    writer.WriteStartElement("runtime");
                    writer.WriteStartElement("assemblyBinding", "urn:schemas-microsoft-com:asm.v1");
                    writer.WriteStartElement("dependentAssembly");

                    writer.WriteStartElement("assemblyIdentity");
                    writer.WriteAttributeString("name", "Newtonsoft.Json");
                    writer.WriteAttributeString("publicKeyToken", "30ad4fe6b2a6aeed");
                    writer.WriteAttributeString("culture", "neutral");
                    writer.WriteEndElement();

                    writer.WriteStartElement("bindingRedirect");
                    writer.WriteAttributeString("oldVersion", "0.0.0.0-11.0.0.0");
                    writer.WriteAttributeString("newVersion", "11.0.0.0");
                    writer.WriteEndElement();

                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("system.diagnostics");
                    writer.WriteStartElement("sharedListeners");
                    writer.WriteStartElement("add");

                    writer.WriteAttributeString("name", "console");
                    writer.WriteAttributeString("type", "System.Diagnostics.ConsoleTraceListener");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("sources");
                    writer.WriteStartElement("source");

                    writer.WriteAttributeString("name", "OfficeDevPnP.Core");
                    writer.WriteAttributeString("switchValue", "Verbose");

                    writer.WriteStartElement("listeners");
                    writer.WriteStartElement("add");
                    writer.WriteAttributeString("name", "console");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("trace");
                    writer.WriteAttributeString("indentsize", "0");
                    writer.WriteAttributeString("autoflush", "true");
                    writer.WriteStartElement("listeners");
                    writer.WriteStartElement("add");
                    writer.WriteAttributeString("name", "console");
                }

            }
        }
        private void WriteProperty(XmlWriter writer, string propertyName, string propertyValue)
        {
            writer.WriteStartElement("add");
            writer.WriteAttributeString("key", propertyName);
            writer.WriteAttributeString("value", propertyValue);
            writer.WriteEndElement();
        }

        private static string GetConnectionString(string c)
        {
            var c2 = c.Substring(c.IndexOf("\"") + 1);
            return c2.Substring(0, c2.IndexOf("\""));
        }
    }
}
