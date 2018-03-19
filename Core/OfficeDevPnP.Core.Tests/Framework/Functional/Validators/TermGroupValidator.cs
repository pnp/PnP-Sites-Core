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

    public class SerializedTermGroupInstance
    {
        public string SchemaXml { get; set; }
    }

    public class TermGroupValidator : ValidatorBase
    {
        #region construction        
        public TermGroupValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:TermGroups/pnp:TermGroup";
        }

        public TermGroupValidator(ClientContext cc) : this()
        {
            this.cc = cc;
        }

        #endregion

        #region Validation logic
        public bool Validate(TermGroupCollection sourceCollection, TermGroupCollection targetCollection, TokenParser tokenParser)
        {
            // Convert object collections to XML 
            List<SerializedTermGroupInstance> sourceTermGroups = new List<SerializedTermGroupInstance>();
            List<SerializedTermGroupInstance> targetTermGroups = new List<SerializedTermGroupInstance>();

            foreach (TermGroup termGroup in sourceCollection)
            {
                ProvisioningTemplate pt = new ProvisioningTemplate();
                pt.TermGroups.Add(termGroup);

                sourceTermGroups.Add(new SerializedTermGroupInstance() { SchemaXml = ExtractElementXml(pt) });
            }

            foreach (TermGroup termGroup in targetCollection)
            {
                ProvisioningTemplate pt = new ProvisioningTemplate();
                pt.TermGroups.Add(termGroup);

                targetTermGroups.Add(new SerializedTermGroupInstance() { SchemaXml = ExtractElementXml(pt) });
            }

            // Use XML validation logic to compare source and target
            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            parserSettings.Add("SchemaXml", null);
            bool isTermGroupsMatch = ValidateObjectsXML(sourceTermGroups, targetTermGroups, "SchemaXml", new List<string> { "Name" }, tokenParser, parserSettings);

            Console.WriteLine("-- Term group validation " + isTermGroupsMatch);
            return isTermGroupsMatch;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            #region TermGroup handling
            // The engine always returns an ID but one does not have to specify an id
            if (sourceObject.Attribute("ID") == null)
            {
                DropAttribute(targetObject, "ID");
            }
            if (sourceObject.Attribute("Description") == null)
            {
                DropAttribute(targetObject, "Description");
            }
            #endregion

            #region TermSet handling
            var sourceTermSets = sourceObject.Descendants(ns + "TermSet");
            var targetTermSets = targetObject.Descendants(ns + "TermSet");
            if (sourceTermSets != null && sourceTermSets.Any())
            {
                foreach (var sourceTermSet in sourceTermSets.ToList())
                {
                    // find relevant target termset
                    var targetTermSet = targetTermSets.FirstOrDefault(p => p.Attribute("Name").Value.Equals(sourceTermSet.Attribute("Name").Value, StringComparison.InvariantCultureIgnoreCase));
                    if (targetTermSet != null)
                    {
                        // The engine always returns an ID (unless for site collection termsets) but one does not have to specify an ID
                        if (sourceTermSet.Attribute("ID") == null || targetTermSet.Attribute("ID") == null)
                        {
                            DropAttribute(targetTermSet, "ID");
                            DropAttribute(sourceTermSet, "ID");
                        }
                        if (sourceTermSet.Attribute("Description") == null)
                        {
                            DropAttribute(targetTermSet, "Description");
                        }
                        if (sourceTermSet.Attribute("Language") == null || targetTermSet.Attribute("Language") == null)
                        {
                            DropAttribute(targetTermSet, "Language");
                            DropAttribute(sourceTermSet, "Language");
                        }

                        #region Term handling
                        var sourceTerms = sourceTermSet.Descendants(ns + "Term");
                        var targetTerms = targetTermSet.Descendants(ns + "Term");

                        foreach (var sourceTerm in sourceTerms.ToList())
                        {
                            // find relevant target term
                            var targetTerm = targetTerms.FirstOrDefault(p => p.Attribute("Name").Value.Equals(sourceTerm.Attribute("Name").Value, StringComparison.InvariantCultureIgnoreCase));
                            if (targetTerm != null)
                            {
                                // The engine always returns an ID (unless for site collection terms) but one does not have to specify an ID
                                if (sourceTerm.Attribute("ID") == null || targetTerm.Attribute("ID") == null)
                                {
                                    DropAttribute(targetTerm, "ID");
                                    DropAttribute(sourceTerm, "ID");
                                }
                                if (sourceTerm.Attribute("CustomSortOrder") == null)
                                {
                                    DropAttribute(targetTerm, "CustomSortOrder");
                                }
                                if (sourceTerm.Attribute("Owner") == null)
                                {
                                    DropAttribute(targetTerm, "Owner");
                                }
                                if (sourceTerm.Attribute("SourceTermId") == null)
                                {
                                    DropAttribute(targetTerm, "SourceTermId");
                                }
                                if (sourceTerm.Attribute("IsSourceTerm") == null)
                                {
                                    DropAttribute(targetTerm, "IsSourceTerm");
                                }
                                if (sourceTerm.Attribute("IsReused") == null)
                                {
                                    DropAttribute(targetTerm, "IsReused");
                                }
                                else
                                {
                                    if ((sourceTerm.Attribute("IsReused").Value.ToBoolean() && 
                                        (sourceTerm.Attribute("IsSourceTerm") != null && !sourceTerm.Attribute("IsSourceTerm").Value.ToBoolean())))
                                    {
                                        // When a reused term had custom local properties defined then we'll export both the local properties of the "base" term and the local properties 
                                        // of reused term. This needs to be taken in account in comparison.
                                        var sourceLocalCustomProperties = sourceTerm.Descendants(ns + "LocalCustomProperties");
                                        if (sourceLocalCustomProperties != null && sourceLocalCustomProperties.Any())
                                        {
                                            var targetLocalCustomProperties = targetTerm.Descendants(ns + "LocalCustomProperties");
                                            if (targetLocalCustomProperties != null && targetLocalCustomProperties.Any())
                                            {
                                                // get target local properties
                                                foreach(var targetLocalCustomProperty in targetLocalCustomProperties.Descendants().ToList())
                                                {
                                                    var sourceLocalCustomProperty = sourceLocalCustomProperties.Descendants().FirstOrDefault(p => p.Attribute("Key").Value.Equals(targetLocalCustomProperty.Attribute("Key").Value, StringComparison.InvariantCultureIgnoreCase));
                                                    if (sourceLocalCustomProperty == null)
                                                    {
                                                        targetLocalCustomProperty.Remove();
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // delete the target terms which were not defined in the source
                        foreach (var targetTerm in targetTerms.ToList())
                        {
                            var sourceTerm = sourceTerms.FirstOrDefault(p => p.Attribute("Name").Value.Equals(targetTerm.Attribute("Name").Value, StringComparison.InvariantCultureIgnoreCase));
                            if (sourceTerm == null)
                            {
                                targetTerm.Remove();
                            }
                        }
                        #endregion
                    }
                }

                // delete the target termsets which were not defined in the source
                foreach (var targetTermSet in targetTermSets.ToList())
                {
                    // find relevant target termset
                    var sourceTermSet = sourceTermSets.FirstOrDefault(p => p.Attribute("Name").Value.Equals(targetTermSet.Attribute("Name").Value, StringComparison.InvariantCultureIgnoreCase));
                    if (sourceTermSet == null)
                    {
                        targetTermSet.Remove();
                    }
                }
            }
            #endregion

        }

        #endregion
    }
}
