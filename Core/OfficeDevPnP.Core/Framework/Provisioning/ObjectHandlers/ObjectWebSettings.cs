using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.IO;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWebSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Web Settings"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(
#if !CLIENTSDKV15
                    w => w.NoCrawl,
                    w => w.RequestAccessEmail,
#endif
                    w => w.MasterUrl,
                    w => w.CustomMasterUrl,
                    w => w.SiteLogoUrl,
                    w => w.RootFolder,
                    w => w.AlternateCssUrl);

                var webSettings = new WebSettings();
#if !CLIENTSDKV15
                webSettings.NoCrawl = web.NoCrawl;
                webSettings.RequestAccessEmail = web.RequestAccessEmail;
#endif
                webSettings.MasterPageUrl = web.MasterUrl;
                webSettings.CustomMasterPageUrl = web.CustomMasterUrl;
                webSettings.SiteLogo = web.SiteLogoUrl;
                webSettings.WelcomePage = web.RootFolder.WelcomePage;
                webSettings.AlternateCSS = web.AlternateCssUrl;
                template.WebSettings = webSettings;

                if (creationInfo.PersistBrandingFiles)
                {
                    if (!string.IsNullOrEmpty(web.MasterUrl))
                    {
                        var masterUrl = web.MasterUrl.ToLower();
                        if (!masterUrl.EndsWith("default.master") && !masterUrl.EndsWith("custom.master") && !masterUrl.EndsWith("v4.master") && !masterUrl.EndsWith("seattle.master") && !masterUrl.EndsWith("oslo.master"))
                        {

                            PersistFile(web, creationInfo, scope, web.MasterUrl);
                        }
                    }
                    if (!string.IsNullOrEmpty(web.CustomMasterUrl))
                    {
                        var customMasterUrl = web.CustomMasterUrl.ToLower();
                        if (!customMasterUrl.EndsWith("default.master") && !customMasterUrl.EndsWith("custom.master") && !customMasterUrl.EndsWith("v4.master") && !customMasterUrl.EndsWith("seattle.master") && !customMasterUrl.EndsWith("oslo.master"))
                        {

                            PersistFile(web, creationInfo, scope, web.CustomMasterUrl);
                        }
                    }
                    if (!string.IsNullOrEmpty(web.SiteLogoUrl))
                    {
                        PersistFile(web, creationInfo, scope, web.SiteLogoUrl);
                    }
                    if (!string.IsNullOrEmpty(web.AlternateCssUrl))
                    {
                        PersistFile(web, creationInfo, scope, web.AlternateCssUrl);
                    }

                }
            }
            return template;
        }

        private void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string serverRelativeUrl)
        {
            if (creationInfo.FileConnector != null)
            {
                try
                {
                    var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
                    string fileName = string.Empty;
                    if (serverRelativeUrl.IndexOf("/") > -1)
                    {
                        fileName = serverRelativeUrl.Substring(serverRelativeUrl.LastIndexOf("/") + 1);
                    }
                    else
                    {
                        fileName = serverRelativeUrl;
                    }
                    web.Context.Load(file);
                    web.Context.ExecuteQueryRetry();
                    ClientResult<Stream> stream = file.OpenBinaryStream();
                    web.Context.ExecuteQueryRetry();

                    using (Stream memStream = new MemoryStream())
                    {
                        CopyStream(stream.Value, memStream);
                        memStream.Position = 0;
                        creationInfo.FileConnector.SaveFileStream(fileName, memStream);
                    }
                }
                catch (ServerException ex1)
                {
                    // If we are referring a file from a location outside of the current web or at a location where we cannot retrieve the file an exception is thrown. We swallow this exception.
                    if (ex1.ServerErrorCode != -2147024809)
                    {
                        throw;
                    }
                    else
                    {
                        scope.LogWarning("File is not necessarily located in the current web. Not retrieving {0}", serverRelativeUrl);
                    }
                }
            }
            else
            {
                WriteWarning("No connector present to persist homepage.", ProvisioningMessageType.Error);
                scope.LogError("No connector present to persist homepage");
            }

        }

        private void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.WebSettings != null)
                {
                    var webSettings = template.WebSettings;
#if !CLIENTSDKV15
                    web.NoCrawl = webSettings.NoCrawl;

                    String requestAccessEmailValue = parser.ParseString(webSettings.RequestAccessEmail);
                    if (!String.IsNullOrEmpty(requestAccessEmailValue) && requestAccessEmailValue.Length >= 255)
                    {
                        requestAccessEmailValue = requestAccessEmailValue.Substring(0, 255);
                    }
                    if (!String.IsNullOrEmpty(requestAccessEmailValue))
                    {
                        web.RequestAccessEmail = requestAccessEmailValue;
                    }
#endif
                    var masterUrl = parser.ParseString(webSettings.MasterPageUrl);
                    if (!string.IsNullOrEmpty(masterUrl))
                    {
                        web.MasterUrl = masterUrl;
                    }
                    var customMasterUrl = parser.ParseString(webSettings.CustomMasterPageUrl);
                    if (!string.IsNullOrEmpty(customMasterUrl))
                    {
                        web.CustomMasterUrl = customMasterUrl;
                    }
                    web.Description = parser.ParseString(webSettings.Description);
                    web.SiteLogoUrl = parser.ParseString(webSettings.SiteLogo);
                    var welcomePage = parser.ParseString(webSettings.WelcomePage);
                    if (!string.IsNullOrEmpty(welcomePage))
                    {
                        web.RootFolder.WelcomePage = welcomePage;
                        web.RootFolder.Update();
                    }
                    web.AlternateCssUrl = parser.ParseString(webSettings.AlternateCSS);

                    web.Update();
                    web.Context.ExecuteQueryRetry();
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return template.WebSettings != null;
        }
    }
}
