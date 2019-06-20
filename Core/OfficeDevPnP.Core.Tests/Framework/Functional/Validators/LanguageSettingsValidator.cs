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

    public class LanguageSettingsValidator : ValidatorBase
    {
        #region construction        
        public LanguageSettingsValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:SupportedUILanguages";
        }

        public LanguageSettingsValidator(ClientContext cc) : this()
        {
            this.cc = cc;
        }

        #endregion

        #region Validation logic
        public bool Validate(SupportedUILanguageCollection sourceLanguageSettings, SupportedUILanguageCollection targetLanguageSettings, TokenParser tokenParser)
        {
            ProvisioningTemplate sourcePt = new ProvisioningTemplate();
            sourcePt.SupportedUILanguages = sourceLanguageSettings;
            var sourceXml = ExtractElementXml(sourcePt);

            ProvisioningTemplate targetPt = new ProvisioningTemplate();
            targetPt.SupportedUILanguages = targetLanguageSettings;
            var targetXml = ExtractElementXml(targetPt);

            return ValidateObjectXML(sourceXml, targetXml, null);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

        }
        #endregion

        #region Helper methods
        #endregion
    }
}
