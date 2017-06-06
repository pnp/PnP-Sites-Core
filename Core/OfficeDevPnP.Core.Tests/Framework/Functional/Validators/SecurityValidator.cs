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
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class SecurityValidator : ValidatorBase
    {
        private TokenParser tParser;
        public class SerializedSecurity
        {
            public string SchemaXml { get; set; }
        }
        public SecurityValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }

        public bool Validate(SiteSecurity source, SiteSecurity target, TokenParser parser, Microsoft.SharePoint.Client.ClientContext context)
        {
            tParser = parser;
            cc = context;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:Security";

            ProvisioningTemplate pt = new ProvisioningTemplate();
            pt.Security = source;
            string sSchemaXml = ExtractElementXml(pt);

            ProvisioningTemplate ptTarget = new ProvisioningTemplate();
            ptTarget.Security = target;
            string tSchemaXml = ExtractElementXml(ptTarget);

            // Use XML validation logic to compare source and target
            if (!ValidateObjectXML(sSchemaXml, tSchemaXml, null)) { return false; }

            return true;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            #region "Group Users"
            AddLoginNameForSourceUsers(sourceObject, "User", "Name");
            RemoveClaimsPrefix(targetObject, "User", "Name");
            DeleteTargetIfNotExistsInSource(sourceObject, targetObject, "User", "Name");
            #endregion

            #region "Site Groups"
            DeleteTargetIfNotExistsInSource(sourceObject, targetObject, "SiteGroup", "Title");
            RemoveClaimsPrefix(targetObject, "SiteGroup", "Owner");
            #endregion

            #region "Permissions - RoleDefinitions"
            DeleteTargetIfNotExistsInSource(sourceObject, targetObject, "RoleDefinition", "Name");
            // remove target permissions if not exists in source
            var sPermissions = sourceObject.Descendants(ns + "Permission");
            if (sPermissions != null && sPermissions.Any())
            {
                var permissions = targetObject.Descendants(ns + "Permission");
                if (permissions != null && permissions.Any())
                {
                    foreach (var permission in permissions.ToList())
                    {
                        if (!sPermissions.Where(p => p.Value == permission.Value).Any())
                        {
                            permission.Remove();
                        }
                    }
                }
            }
            #endregion

            #region "Permissions - RoleAssignments"
            ParseStringsInRoleAssignment(targetObject, "RoleAssignment");
            RemoveClaimsPrefix(targetObject, "RoleAssignment", "Principal");
            // remove target role assignments if not exists in source
            var sColl = sourceObject.Descendants(ns + "RoleAssignment");
            if (sColl != null && sColl.Any())
            {
                var tColl = targetObject.Descendants(ns + "RoleAssignment");
                if (tColl != null && tColl.Any())
                {
                    foreach (var item in tColl.ToList())
                    {
                        if (!sColl.Where(i => i.Attribute("Principal").Value == item.Attribute("Principal").Value && i.Attribute("RoleDefinition").Value == item.Attribute("RoleDefinition").Value).Any())
                        {
                            item.Remove();
                        }
                    }
                }
            }
            #endregion
        }

        private void ParseStringsInRoleAssignment(XElement targetObject, string elementName)
        {
            XNamespace ns = SchemaVersion;
            IEnumerable<XElement> coll = targetObject.Descendants(ns + elementName);
            foreach (var item in coll)
            {
                item.Attribute("Principal").Value = tParser.ParseString(item.Attribute("Principal").Value);
                item.Attribute("RoleDefinition").Value = tParser.ParseString(item.Attribute("RoleDefinition").Value);
            }

        }

        private void RemoveClaimsPrefix(XElement targetObject, string elementName, string attributeName)
        {
            XNamespace ns = SchemaVersion;
            IEnumerable<XElement> coll = targetObject.Descendants(ns + elementName);
            string name = "";

            foreach (var item in coll)
            {
                name = item.Attribute(attributeName).Value;
                if (name.Contains("|") && !name.StartsWith("c:"))
                {
                    item.Attribute(attributeName).Value = name.Substring(name.LastIndexOf(("|")) + 1);
                }
            }
        }
        private void AddLoginNameForSourceUsers(XElement sourceObject, string elementName, string attributeName)
        {
            XNamespace ns = SchemaVersion;
            IEnumerable<XElement> coll = sourceObject.Descendants(ns + elementName);
            string name = "";
            string loginName = "";
            foreach (var item in coll)
            {
                name = item.Attribute(attributeName).Value;
                if (!name.Contains("@"))
                {
                    var existingUser =cc.Web.EnsureUser(name);
                    cc.Web.Context.Load(existingUser);
                    cc.Web.Context.ExecuteQueryRetry();
                    loginName = existingUser.LoginName;
                    if (loginName.Contains("|") && !name.StartsWith("c:"))
                    {
                        loginName= loginName.Substring(loginName.LastIndexOf(("|")) + 1);
                    }
                    item.Attribute(attributeName).Value = loginName;
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

    }
}
