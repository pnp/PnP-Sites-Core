using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{

    public class WebSettingsValidator : ValidatorBase
    {
        private bool isNoScriptSite = false;

        #region construction        
        public WebSettingsValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:WebSettings";
        }

        public WebSettingsValidator(ClientContext cc) : this()
        {
            this.cc = cc;
            isNoScriptSite = cc.Web.IsNoScriptSite();
        }

        #endregion

        #region Validation logic
        public bool Validate(WebSettings sourceWebsettings, WebSettings targetWebSettings, TokenParser tokenParser)
        {
            ProvisioningTemplate sourcePt = new ProvisioningTemplate();
            sourcePt.WebSettings = sourceWebsettings;
            var sourceXml = ExtractElementXml(sourcePt);

            ProvisioningTemplate targetPt = new ProvisioningTemplate();
            targetPt.WebSettings = targetWebSettings;
            var targetXml = ExtractElementXml(targetPt);

            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            List<string> parsedProperties = new List<string>();
            parsedProperties.AddRange(new string[] { "MasterPageUrl", "CustomMasterPageUrl", "Description", "Title", "SiteLogo", "WelcomePage", "AlternateCSS" });

            return ValidateObjectXML(sourceXml, targetXml, parsedProperties, tokenParser, parserSettings);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            // the engine is not extracting title and description, only allows to set them
            DropAttribute(sourceObject, "Title");
            DropAttribute(sourceObject, "Description");

            // master pages are extracted relative to the root site without token...e.g. /_catalogs/MasterPage/oslo.master.
            // given we can use tokens in the template we do a manual comparison and drop the MasterPageUrl and CustomMasterPageUrl attributes when ok
            if (isNoScriptSite ||
                (ValidateMasterPage(sourceObject.Attribute("MasterPageUrl").Value, targetObject.Attribute("MasterPageUrl").Value) &&
                ValidateMasterPage(sourceObject.Attribute("CustomMasterPageUrl").Value, targetObject.Attribute("CustomMasterPageUrl").Value)))
            {
                DropAttribute(sourceObject, "MasterPageUrl");
                DropAttribute(sourceObject, "CustomMasterPageUrl");
                DropAttribute(targetObject, "MasterPageUrl");
                DropAttribute(targetObject, "CustomMasterPageUrl");
            }

#if ONPREMISES
            // we don't support NoCrawl in on-premises so drop them from source and target
            DropAttribute(sourceObject, "NoCrawl");
            DropAttribute(targetObject, "NoCrawl");
#endif

            if (isNoScriptSite)
            {
                DropAttribute(sourceObject, "NoCrawl");
                DropAttribute(targetObject, "NoCrawl");
            }

        }
        #endregion

        #region Helper methods
        private bool ValidateMasterPage(string source, string target)
        {
            if (!source.StartsWith("/_catalogs/MasterPage", StringComparison.InvariantCultureIgnoreCase))
            {
                int loc = source.IndexOf("/_catalogs");
                if (loc >= 0)
                {
                    source = source.Substring(loc);

                    if (!source.Equals(target, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return false;
                    }
                }
            }

            return true;
        }
#endregion
    }
}
