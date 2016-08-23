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

    public class RegionalSettingsValidator : ValidatorBase
    {
        #region construction        
        public RegionalSettingsValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:RegionalSettings";
        }

        public RegionalSettingsValidator(ClientContext cc) : this()
        {
            this.cc = cc;
        }

        #endregion

        #region Validation logic
        public bool Validate(Core.Framework.Provisioning.Model.RegionalSettings sourceRegionalSettings, Core.Framework.Provisioning.Model.RegionalSettings targetRegionalSettings, TokenParser tokenParser)
        {
            ProvisioningTemplate sourcePt = new ProvisioningTemplate();
            sourcePt.RegionalSettings = sourceRegionalSettings;
            var sourceXml = ExtractElementXml(sourcePt);

            ProvisioningTemplate targetPt = new ProvisioningTemplate();
            targetPt.RegionalSettings = targetRegionalSettings;
            var targetXml = ExtractElementXml(targetPt);

            return ValidateObjectXML(sourceXml, targetXml, null, null, null);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            if (sourceObject.Attribute("CalendarType").Value.ToLower() == "none")
            {
                DropAttribute(sourceObject, "CalendarType");
                DropAttribute(targetObject, "CalendarType");
            }

        }

        #endregion
    }
}
