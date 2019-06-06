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
    public class WorkflowValidator : ValidatorBase
    {
        private TokenParser tParser;
        public class SerializedSecurity
        {
            public string SchemaXml { get; set; }
        }
        public WorkflowValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }

        public bool Validate(Workflows source, Workflows target, TokenParser parser)
        {
            tParser = parser;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:Workflows";

            ProvisioningTemplate pt = new ProvisioningTemplate();
            pt.Workflows = source;
            string sSchemaXml = ExtractElementXml(pt);

            ProvisioningTemplate ptTarget = new ProvisioningTemplate();
            ptTarget.Workflows = target;
            string tSchemaXml = ExtractElementXml(ptTarget);

            // Use XML validation logic to compare source and target
            if (!ValidateObjectXML(sSchemaXml, tSchemaXml, null)) { return false; }

            return true;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            #region "Workflow definitions"
            DeleteTargetIfNotExistsInSource(sourceObject, targetObject, "WorkflowDefinition", "Id");
            #endregion

            #region "properties"
            // Properties are too different, no point in trying to validate as they're also not an indication of engine WF deployment success or not
            var WFproperties = sourceObject.Descendants(ns + "Properties");
            if (WFproperties != null)
            {
                WFproperties.Remove();
            }
            WFproperties = targetObject.Descendants(ns + "Properties");
            if (WFproperties != null)
            {
                WFproperties.Remove();
            }
            var WFPropertyDefinitions = sourceObject.Descendants(ns + "PropertyDefinitions");
            if (WFPropertyDefinitions != null)
            {
                WFPropertyDefinitions.Remove();
            }
            WFPropertyDefinitions = targetObject.Descendants(ns + "PropertyDefinitions");
            if (WFPropertyDefinitions != null)
            {
                WFPropertyDefinitions.Remove();
            }

            RemoveAttribute(sourceObject, "WorkflowSubscription", "StatusFieldName");
            RemoveAttribute(targetObject, "WorkflowSubscription", "StatusFieldName");

            #endregion

            #region removing xaml path attribute 
            RemoveAttribute(targetObject, "WorkflowDefinition", "XamlPath");
            RemoveAttribute(sourceObject, "WorkflowDefinition", "XamlPath");
            #endregion

            #region "Property Tag"
            // remove target Property Tag if not exists in source
            var sPproperties = sourceObject.Descendants(ns + "Properties");
            if (sPproperties != null && !sPproperties.Any())
            {
                var tProperties = targetObject.Descendants(ns + "Properties");
                if (tProperties != null && tProperties.Any())
                {
                    tProperties.Remove();
                }
            }
            #endregion


        }
        private void RemoveAttribute(XElement targetObject, string elementName, string attributeName)
        {
            XNamespace ns = SchemaVersion;
            IEnumerable<XElement> coll = targetObject.Descendants(ns + elementName);
            string name = "";
            foreach (var item in coll)
            {
                name = item.Attribute(attributeName).Value;
                if (name!=null)
                {
                    item.Attribute(attributeName).Remove();
                    break;
                }
            }
        }
        private void DeleteTargetIfNotExistsInSource(XElement sourceObject, XElement targetObject, string elementName, string key)
        {
            XNamespace ns = SchemaVersion;

            var sColl = sourceObject.Descendants(ns + elementName);
            if (sColl != null && sColl.Any())
            {
                var tColl = targetObject.Descendants(ns + elementName);
                if (tColl != null && tColl.Any())
                {
                    foreach (var item in tColl.ToList())
                    {
                        if (!sColl.Where(u => u.Attribute(key).Value == item.Attribute(key).Value).Any())
                        {
                            item.Remove();
                        }
                    }
                }
            }
        }

        private void DeleteTargetIfNotExistsInSource_Properties(XElement sourceObject, XElement targetObject, string elementName, string key)
        {
            XNamespace ns = SchemaVersion;

            var sColl = sourceObject.Descendants(ns + elementName);
            var tColl = targetObject.Descendants(ns + elementName);

            if ((sColl != null && sColl.Any())  || !sColl.Any()&&tColl!=null&&tColl.Any() )
            {
                if (tColl != null && tColl.Any())
                {
                    foreach (var item in tColl.ToList())
                    {
                        if (!sColl.Where(u => u.Attribute(key).Value == item.Attribute(key).Value).Any())
                        {
                            item.Remove();
                        }
                    }
                }
            }
            
        }

    }
}
