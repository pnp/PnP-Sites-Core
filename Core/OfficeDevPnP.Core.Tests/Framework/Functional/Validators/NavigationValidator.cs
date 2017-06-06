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
    /// <summary>
    /// Summary description for NavigationValidator
    /// </summary>
    [TestClass]
    public class NavigationValidator : ValidatorBase
    {
        public class SerializedNavigation
        {
            public string SchemaXml { get; set; }
        }

        private TokenParser navigationParser;

        public NavigationValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:Navigation";

        }

        public bool Validate(Navigation source, Navigation target, TokenParser tokenParser)
        {
            ProvisioningTemplate ptSource = new ProvisioningTemplate();
            ptSource.Navigation = source;
            var sourceXml = ExtractElementXml(ptSource);

            ProvisioningTemplate ptTarget = new ProvisioningTemplate();
            ptTarget.Navigation = target;
            var targetXml = ExtractElementXml(ptTarget);

            navigationParser = tokenParser;

            return ValidateObjectXML(sourceXml, targetXml, null);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;
            var firstStructuralNavTarget = targetObject.Descendants(ns + "StructuralNavigation").FirstOrDefault();
            var lastStructuralNavTarget = targetObject.Descendants(ns + "StructuralNavigation").LastOrDefault();
            var firstStructuralNavSource = sourceObject.Descendants(ns + "StructuralNavigation").FirstOrDefault();
            var lastStructuralNavSource = sourceObject.Descendants(ns + "StructuralNavigation").LastOrDefault();

            // Drop the RemoveExistingNodes attribute before the comparison. Retrieved target xml contains "RemoveExistingNodes=false" value.
            if (firstStructuralNavTarget != null && firstStructuralNavSource != null)
            {
                firstStructuralNavTarget.Attribute("RemoveExistingNodes").Remove();
                firstStructuralNavSource.Attribute("RemoveExistingNodes").Remove();
            }

            if (lastStructuralNavTarget != null && targetObject.Descendants(ns + "StructuralNavigation").Count() > 1 && sourceObject.Descendants(ns + "StructuralNavigation").Count() > 1)
            {
                lastStructuralNavTarget.Attribute("RemoveExistingNodes").Remove();
                lastStructuralNavSource.Attribute("RemoveExistingNodes").Remove();
            }

            // Drop the NavigationType attribute before the comparison. In subsite, 'StructuralLocal' navigation type is retrieved when 'Structural' navigation type provisioned.
            var currentNavSource = sourceObject.Descendants(ns + "CurrentNavigation").FirstOrDefault();
            var currentNavTarget = targetObject.Descendants(ns + "CurrentNavigation").FirstOrDefault();
            if (currentNavSource.Attribute("NavigationType").Value == "Structural" && currentNavTarget.Attribute("NavigationType").Value == "StructuralLocal") // Structural != StructuralLocal
            {
                currentNavSource.Attribute("NavigationType").Remove();
                currentNavTarget.Attribute("NavigationType").Remove();
            }

            var managedSource = sourceObject.Descendants(ns + "ManagedNavigation").FirstOrDefault();
            if (managedSource != null && managedSource.Attribute("TermSetId") != null)
            {
                managedSource.Attribute("TermSetId").Value = navigationParser.ParseString(managedSource.Attribute("TermSetId").Value);
                var managedTarget = targetObject.Descendants(ns + "ManagedNavigation").FirstOrDefault();
                if (managedTarget != null && managedTarget.Attribute("TermSetId") != null)
                {
                    managedTarget.Attribute("TermSetId").Value = navigationParser.ParseString(managedTarget.Attribute("TermSetId").Value);
                }
            }
        }
    }
}
