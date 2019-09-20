using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    #region Delegates
    public delegate void ValidateEventHandler(object sender, ValidateEventArgs e);
    public delegate void ValidateXmlEventHandler(object sender, ValidateXmlEventArgs e);
    #endregion

    /// <summary>
    /// Base object validator class
    /// </summary>
    public class ValidatorBase
    {
        #region Variables
        internal ClientContext cc = null;
        #endregion

        #region Events
        public event ValidateEventHandler ValidateEvent;
        public event ValidateXmlEventHandler ValidateXmlEvent;
        #endregion

        #region public properties
        public string SchemaVersion { get; set; }
        public string XPathQuery { get; set; }
        #endregion

        #region construction
        public ValidatorBase()
        {
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
        }
        #endregion

        #region Validation methods

        public virtual bool ValidateObjects<T>(T sourceElement, T targetElement, List<string> properties) where T : class
        {
            IEnumerable sElements = (IEnumerable)sourceElement;
            IEnumerable tElements = (IEnumerable)targetElement;

            string key = properties[0];
            int sourceCount = 0;
            int targetCount = 0;

            foreach (string property in properties)
            {
                foreach (object sElem in sElements)
                {
                    sourceCount++;
                    string sourceKey = sElem.GetType().GetProperty(property).GetValue(sElem).ToString();

                    foreach (object tElem in tElements)
                    {
                        string targetKey = tElem.GetType().GetProperty(property).GetValue(tElem).ToString();

                        if (sourceKey.Equals(targetKey))
                        {
                            targetCount++;
                            break;
                        }
                    }
                }
            }

            return sourceCount == targetCount;
        }

        public virtual bool ValidateObjects<T>(T sourceElement, T targetElement, List<string> properties, TokenParser tokenParser=null, Dictionary<string, string[]> parsedProperties=null) where T : class
        {
            IEnumerable sElements = (IEnumerable)sourceElement;
            IEnumerable tElements = (IEnumerable)targetElement;

            string key = properties[0];
            int sourceCount = 0;
            int targetCount = 0;
            foreach (object sElem in sElements)
            {
                sourceCount++;
                string sourceKey = sElem.GetType().GetProperty(key).GetValue(sElem).ToString();

                if (tokenParser != null && parsedProperties != null)
                {
                    if (parsedProperties.ContainsKey(key))
                    {
                        string[] parserExceptions;
                        parsedProperties.TryGetValue(key, out parserExceptions);
                        sourceKey = tokenParser.ParseString(Convert.ToString(sourceKey), parserExceptions);
                    }
                }

                foreach (object tElem in tElements)
                {
                    string targetKey = tElem.GetType().GetProperty(key).GetValue(tElem).ToString();

                    if (sourceKey.Equals(targetKey))
                    {
                        targetCount++;
                        //compare objects
                        foreach(string property in properties)
                        {
                            string sourceProperty = sElem.GetType().GetProperty(property).GetValue(sElem).ToString();
                            if (tokenParser != null && parsedProperties != null)
                            {
                                if (parsedProperties.ContainsKey(property))
                                {
                                    string[] parserExceptions;
                                    parsedProperties.TryGetValue(key, out parserExceptions);
                                    sourceProperty = tokenParser.ParseString(Convert.ToString(sourceProperty), parserExceptions);
                                }
                            }

                            string targetProperty = tElem.GetType().GetProperty(property).GetValue(tElem).ToString();

                            ValidateEventArgs e = null;
                            if (ValidateEvent != null)
                            {
                                e = new ValidateEventArgs(property, sourceProperty, targetProperty, sElem, tElem);
                                ValidateEvent(this, e);
                            }

                            if (e != null && e.IsEqual)
                            {
                                // Do nothing since we've declared equality in the event handler
                            }
                            else
                            {
                                if (!sourceProperty.Equals(targetProperty))
                                {
                                    return false;
                                }
                            }
                        }
                        break;
                    }
                }
            }

            return sourceCount == targetCount;
        }

        public virtual bool ValidateObjectsXML<T>(IEnumerable<T> sElements, IEnumerable<T> tElements, string XmlPropertyName, List<string> properties, TokenParser tokenParser = null, Dictionary<string, string[]> parsedProperties = null) where T: class
        {
            string key = properties[0];
            int sourceCount = 0;
            int targetCount = 0;

            foreach (var sElem in sElements)
            {
                sourceCount++;
                string sourceXmlString = sElem.GetType().GetProperty(XmlPropertyName).GetValue(sElem).ToString();

                if (tokenParser != null && parsedProperties != null)
                {
                    if (parsedProperties.ContainsKey(XmlPropertyName))
                    {
                        string[] parserExceptions;
                        parsedProperties.TryGetValue(XmlPropertyName, out parserExceptions);
                        sourceXmlString = tokenParser.ParseString(Convert.ToString(sourceXmlString), parserExceptions);
                    }
                }

                XElement sourceXml = XElement.Parse(sourceXmlString);
                string sourceKeyValue = sourceXml.Attribute(key).Value;

                foreach (var tElem in tElements)
                {
                    string targetXmlString = tElem.GetType().GetProperty(XmlPropertyName).GetValue(tElem).ToString();

                    if (tokenParser != null && parsedProperties != null)
                    {
                        if (parsedProperties.ContainsKey(XmlPropertyName))
                        {
                            string[] parserExceptions;
                            parsedProperties.TryGetValue(XmlPropertyName, out parserExceptions);
                            targetXmlString = tokenParser.ParseString(Convert.ToString(targetXmlString), parserExceptions);
                        }
                    }

                    XElement targetXml = XElement.Parse(targetXmlString);
                    string targetKeyValue = targetXml.Attribute(key).Value;

                    if (sourceKeyValue.Equals(targetKeyValue, StringComparison.InvariantCultureIgnoreCase))
                    {
                        targetCount++;

                        // compare XML's

                        // call virtual override method, consuming validators can add fixed validation logic if needed
                        OverrideXmlData(sourceXml, targetXml);

                        // call event handler, validator instances can add additional validation behaviour if needed, including forcing an IsEqual
                        ValidateXmlEventArgs e = null;
                        if (ValidateXmlEvent != null)
                        {
                            e = new ValidateXmlEventArgs(sourceXml, targetXml);
                            ValidateXmlEvent(this, e);
                        }

                        if (e != null && e.IsEqual)
                        {
                            // Do nothing since we've declared equality in the event handler
                        }
                        else
                        {
                            // Not using XNode.DeepEquals anymore since it requires that the attributes in both XML's are ordered the same
                            //var equalNodes = XNode.DeepEquals(sourceXml, targetXml);
                            var equalNodes = XmlComparer.AreEqual(sourceXml, targetXml);
                            if (!equalNodes.Success)
                            {
                                Console.WriteLine(string.Format("Source XML:{0}", sourceXml.ToString()));
                                Console.WriteLine(string.Format("Target XML:{0}", targetXml.ToString()));
                                return false;
                            }
                        }

                        break;
                    }
                }
            }

            return sourceCount == targetCount;
        }

        public virtual bool ValidateObjectXML(string source, string target, List<string> properties, TokenParser tokenParser = null, Dictionary<string, string[]> parsedProperties = null) 
        {
            XElement sourceXml = XElement.Parse(source);
            XElement targetXml = XElement.Parse(target);

            if (tokenParser != null && parsedProperties != null)
            {
                // Run token parser over provided list of properties
                foreach (var property in properties)
                {
                    if (sourceXml.Attribute(property) != null)
                    {
                        string[] parserExceptions;
                        parsedProperties.TryGetValue(property, out parserExceptions);
                        sourceXml.Attribute(property).Value = tokenParser.ParseString(sourceXml.Attribute(property).Value, parserExceptions);
                        if (targetXml.Attribute(property) != null)
                        {
                            targetXml.Attribute(property).Value = tokenParser.ParseString(targetXml.Attribute(property).Value, parserExceptions);
                        }
                    }
                }
            }

            // call virtual override method, consuming validators can add fixed validation logic if needed
            OverrideXmlData(sourceXml, targetXml);

            // call event handler, validator instances can add additional validation behaviour if needed, including forcing an IsEqual
            ValidateXmlEventArgs e = null;
            if (ValidateXmlEvent != null)
            {
                e = new ValidateXmlEventArgs(sourceXml, targetXml);
                ValidateXmlEvent(this, e);
            }

            if (e != null && e.IsEqual)
            {
                // Do nothing since we've declared equality in the event handler
            }
            else
            {
                // Not using XNode.DeepEquals anymore since it requires that the attributes in both XML's are ordered the same
                //var equalNodes = XNode.DeepEquals(sourceXml, targetXml);
                var equalNodes = XmlComparer.AreEqual(sourceXml, targetXml);
                if (!equalNodes.Success)
                {
                    Console.WriteLine(string.Format("Source XML:{0}", sourceXml.ToString()));
                    Console.WriteLine(string.Format("Target XML:{0}", targetXml.ToString()));
                    return false;
                }
            }

            return true;
        }

        internal virtual void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {

        }

        internal string ExtractElementXml(ProvisioningTemplate provisioningTemplate)
        {
            XElement provXml = XElement.Parse(provisioningTemplate.ToXML(XMLPnPSchemaFormatter.GetSpecificFormatter(SchemaVersion)));
            var namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace("pnp", SchemaVersion);
            XElement ctXml = provXml.XPathSelectElement(XPathQuery, namespaceManager);
            return ctXml.ToString(SaveOptions.DisableFormatting);
        }

        internal bool ValidateSecurityXml(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            var sourceBreakRoleInheritance = sourceObject.Descendants(ns + "BreakRoleInheritance");
            if (sourceBreakRoleInheritance != null && sourceBreakRoleInheritance.Any())
            {
                bool copyRoleAssignments = sourceBreakRoleInheritance.First().Attribute("CopyRoleAssignments").Value.ToBoolean();

                var sourceRoleAssignments = sourceObject.Descendants(ns + "RoleAssignment");
                var targetRoleAssignments = targetObject.Descendants(ns + "RoleAssignment");

                // Verify the number of role assignments
                if (sourceRoleAssignments.Count() > targetRoleAssignments.Count())
                {
                    return false;
                }

                // verify the added ones appear in target
                foreach (var sourceRoleAssignment in sourceRoleAssignments)
                {
                    var targetRoleAssignment = targetRoleAssignments.Where(p => p.Attribute("Principal").Value == sourceRoleAssignment.Attribute("Principal").Value);

                    if (targetRoleAssignment != null && targetRoleAssignment.Any())
                    {
                        if (!sourceRoleAssignment.Attribute("RoleDefinition").Value.Equals(targetRoleAssignment.First().Attribute("RoleDefinition").Value, StringComparison.InvariantCultureIgnoreCase))
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Verify that the target securable item has the needed security settings
        /// </summary>
        /// <param name="context">Client context object</param>
        /// <param name="security">Security template model</param>
        /// <param name="item">Securable item</param>
        /// <returns></returns>
        public bool ValidateSecurityCSOM(ClientContext context, ObjectSecurity security, SecurableObject item)
        {
            int roleAssignmentCount = security.RoleAssignments.Count;
            int roleCount = 0;

            IEnumerable roles = context.LoadQuery(item.RoleAssignments.Include(roleAsg => roleAsg.Member,
                roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name)));
            context.ExecuteQueryRetry();

            foreach (var s in security.RoleAssignments)
            {
                foreach (Microsoft.SharePoint.Client.RoleAssignment r in roles)
                {
                    if (r.Member.LoginName.Contains(s.Principal))
                    {
                        if (r.Member.LoginName.Equals("c:0(.s|true", StringComparison.InvariantCultureIgnoreCase))
                        {
                            roleCount++;
                        }
                        else if (r.RoleDefinitionBindings.Where(i => i.Name == s.RoleDefinition).FirstOrDefault() != null)
                        {
                            roleCount++;
                        }
                    }
                }
            }

            if (roleAssignmentCount != roleCount)
            {
                return false;
            }

            return true;
        }

        public bool ValidateDataRowsCSOM(ClientContext context, DataRow dataRow, ListItem item)
        {
            context.Load(item);
            context.ExecuteQueryRetry();

            // Validate item values
            foreach(var dataValue in dataRow.Values)
            {
                if (!item[dataValue.Key].ToString().Equals(dataValue.Value, StringComparison.InvariantCultureIgnoreCase))
                {
                    return false;
                }
            }

            // Validate security
            if (dataRow.Security != null)
            {
                if (!ValidateSecurityCSOM(context, dataRow.Security, item))
                {
                    return false;
                }
            }

            return true;
        }

        #endregion

        #region Helper methods
        internal void DropAttribute(XElement xml, string Attribute)
        {
            if (xml.Attribute(Attribute) != null)
            {
                xml.Attribute(Attribute).Remove();
            }
        }

        internal void UpperCaseAttribute(XElement xml, string Attribute)
        {
            if (xml.Attribute(Attribute) != null)
            {
                xml.SetAttributeValue(Attribute, xml.Attribute(Attribute).Value.ToUpper());
            }
        }

        internal bool TaxonomyFieldCustomizationPropertyIsEqual(XElement sourceXml, XElement targetXml, string property)
        {

            //  <Field ID="{35B749BF-0FE3-48F9-A84B-C5EA05246DEB}" Type="TaxonomyFieldType" Name="FLD_50" StaticName="FLD_50" DisplayName="Fld 50" Group="PnP Demo" ShowField="Term1033" Required="FALSE" EnforceUniqueValues="FALSE">
            //  <Customization>
            //    <ArrayOfProperty>
            //      <Property>
            //        <Name>SspId</Name>
            //        <Value xmlns:q1="http://www.w3.org/2001/XMLSchema" p4:type="q1:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">{sitecollectiontermstoreid}</Value>
            //      </Property>
            //      <Property>
            //        <Name>TermSetId</Name>
            //        <Value xmlns:q2="http://www.w3.org/2001/XMLSchema" p4:type="q2:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">{termsetid:TG_1:TS_1}</Value>
            //      </Property>
            //      <Property>
            //        <Name>TextField</Name>
            //        <Value xmlns:q6="http://www.w3.org/2001/XMLSchema" p4:type="q6:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">39E95FAA-894F-4FED-879D-A1A6A8381149</Value>
            //      </Property>
            //      <Property>
            //        <Name>IsPathRendered</Name>
            //        <Value xmlns:q7="http://www.w3.org/2001/XMLSchema" p4:type="q7:boolean" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">false</Value>
            //      </Property>
            //      <Property>
            //        <Name>IsKeyword</Name>
            //        <Value xmlns:q8="http://www.w3.org/2001/XMLSchema" p4:type="q8:boolean" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">false</Value>
            //      </Property>
            //    </ArrayOfProperty>
            //  </Customization>
            //</Field>

            var sourceCustomizationProperty = sourceXml.XPathSelectElement(String.Format("./Customization/ArrayOfProperty/Property[Name = '{0}']/Value", property));
            if (sourceCustomizationProperty != null)
            {
                var targetCustomizationProperty = targetXml.XPathSelectElement(String.Format("./Customization/ArrayOfProperty/Property[Name = '{0}']/Value", property));
                if (targetCustomizationProperty == null)
                {
                    // the property is not present which should never happen
                    return false;
                }
                else
                {
                    // compare property values
                    if (sourceCustomizationProperty.Value.Equals(targetCustomizationProperty.Value, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            else
            {
                // Property not in source can't make comparison fail
                return true;
            }
        }

        #endregion


    }
}
