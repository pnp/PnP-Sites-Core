using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Xml;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class SitePolicyValidator : ValidatorBase
    {
        public SitePolicyValidator() : base()
        {
            // optionally override schema version
            //SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }

        public bool Validate(string source, string target, TokenParser parser)
        {
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:SitePolicy";

            ProvisioningTemplate pt = new ProvisioningTemplate();
            pt.SitePolicy = source;
            string sSchemaXml = ExtractElementXml(pt);

            ProvisioningTemplate ptTarget = new ProvisioningTemplate();
            ptTarget.SitePolicy = target;
            string tSchemaXml = ExtractElementXml(ptTarget);

            // Use XML validation logic to compare source and target
            if (!ValidateObjectXML(sSchemaXml, tSchemaXml, null)) { return false; }

            return true;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;
        }


    }
}
