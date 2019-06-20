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

    public class AuditSettingsValidator : ValidatorBase
    {
        private bool isNoScriptSite = false;

        #region construction        
        public AuditSettingsValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:AuditSettings";
        }

        public AuditSettingsValidator(ClientContext cc) : this()
        {
            this.cc = cc;
            isNoScriptSite = cc.Web.IsNoScriptSite();
        }

        #endregion

        #region Validation logic
        public bool Validate(AuditSettings sourceAuditsettings, AuditSettings targetAuditSettings, TokenParser tokenParser)
        {
            ProvisioningTemplate sourcePt = new ProvisioningTemplate();
            sourcePt.AuditSettings = sourceAuditsettings;
            var sourceXml = ExtractElementXml(sourcePt);

            ProvisioningTemplate targetPt = new ProvisioningTemplate();
            targetPt.AuditSettings = targetAuditSettings;
            var targetXml = ExtractElementXml(targetPt);

            return ValidateObjectXML(sourceXml, targetXml, null);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            if (isNoScriptSite)
            {
                DropAttribute(sourceObject, "AuditLogTrimmingRetention");
                DropAttribute(targetObject, "AuditLogTrimmingRetention");
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
