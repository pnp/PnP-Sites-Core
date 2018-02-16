using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPropertyBagEntry : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Property bag entries"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var systemPropertyBagEntriesExclusions = new List<string>(new[]
                {
                    "_",
                    "vti_",
                    "dlc_",
                    "ecm_",
                    "profileschemaversion",
                    "DesignPreview"
                });

                // Check if this is not a noscript site as we're not allowed to write to the web property bag is that one
                bool isNoScriptSite = web.IsNoScriptSite();
                if (isNoScriptSite)
                {
                    return parser;
                }

                // To handle situations where the propertybag is not updated fully when applying a theme, 
                // we need to create a new context and use that one. Reloading the propertybag does not solve this.
                var webUrl = web.EnsureProperty(w => w.Url);
                var newContext = web.Context.Clone(webUrl);

                web = newContext.Web;

                foreach (var propbagEntry in template.PropertyBagEntries)
                {
                    bool propExists = web.PropertyBagContainsKey(propbagEntry.Key);

                    if (propbagEntry.Overwrite)
                    {
                        var systemProp = systemPropertyBagEntriesExclusions.Any(k => propbagEntry.Key.StartsWith(k, StringComparison.OrdinalIgnoreCase));
                        if (!systemProp || (systemProp && applyingInformation.OverwriteSystemPropertyBagValues))
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_PropertyBagEntries_Overwriting_existing_propertybag_entry__0__with_value__1_, propbagEntry.Key, propbagEntry.Value);
                            web.SetPropertyBagValue(propbagEntry.Key, parser.ParseString(propbagEntry.Value));
                            if (propbagEntry.Indexed)
                            {
                                web.AddIndexedPropertyBagKey(propbagEntry.Key);
                            }
                        }
                    }
                    else
                    {
                        if (!propExists)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_PropertyBagEntries_Creating_new_propertybag_entry__0__with_value__1__2_, propbagEntry.Key, propbagEntry.Value, propbagEntry.Indexed ? ",Indexed = true" : "");
                            web.SetPropertyBagValue(propbagEntry.Key, parser.ParseString(propbagEntry.Value));
                            if (propbagEntry.Indexed)
                            {
                                web.AddIndexedPropertyBagKey(propbagEntry.Key);
                            }
                        }

                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.Context.Load(web, w => w.AllProperties, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();

                var entries = new List<PropertyBagEntry>();

                var indexedProperties = web.GetIndexedPropertyBagKeys().ToList();
                foreach (var propbagEntry in web.AllProperties.FieldValues)
                {
                    var indexed = indexedProperties.Contains(propbagEntry.Key);
                    entries.Add(new PropertyBagEntry() { Key = propbagEntry.Key, Value = propbagEntry.Value.ToString(), Indexed = indexed });
                }

                template.PropertyBagEntries.Clear();
                template.PropertyBagEntries.AddRange(entries);

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo);
                }

                foreach (PropertyBagEntry propbagEntry in template.PropertyBagEntries)
                {
                    propbagEntry.Value = Tokenize(propbagEntry.Value, web.ServerRelativeUrl);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            ProvisioningTemplate baseTemplate = creationInfo.BaseTemplate;

            foreach (var propertyBagEntry in baseTemplate.PropertyBagEntries)
            {
                int index = template.PropertyBagEntries.FindIndex(f => f.Key.Equals(propertyBagEntry.Key));

                if (index > -1)
                {
                    template.PropertyBagEntries.RemoveAt(index);
                }
            }

            // Scan for "system" properties that should be removed as well. Below list contains
            // prefixes of properties that will be dropped
            List<string> systemPropertyBagEntriesExclusions = new List<string>(new string[]
            {
                "_",
                "vti_",
                "dlc_",
                "ecm_",
                "profileschemaversion",
                "DesignPreview"
            });

            // Below property prefixes indicate properties that never can be dropped 
            List<string> systemPropertyBagEntriesInclusions = new List<string>(new string[]
            {
                "_PnP_"
            });
            systemPropertyBagEntriesInclusions.AddRange(creationInfo.PropertyBagPropertiesToPreserve);

            var entriesToDelete = new List<PropertyBagEntry>();

            // Prepare the list of property bag entries that will be dropped
            foreach (var property in systemPropertyBagEntriesExclusions)
            {
                var results = from prop in template.PropertyBagEntries
                              where prop.Key.StartsWith(property, StringComparison.OrdinalIgnoreCase)
                              select prop;
                entriesToDelete.AddRange(results);
            }

            // Remove the property bag entries that we want to forcifully keep
            foreach (var property in systemPropertyBagEntriesInclusions)
            {
                var results = from prop in entriesToDelete
                              where prop.Key.StartsWith(property, StringComparison.OrdinalIgnoreCase)
                              select prop;
                // Drop the found elements from the delete list    
                entriesToDelete = new List<PropertyBagEntry>(SymmetricExcept<PropertyBagEntry>(results, entriesToDelete));
            }

            // Delete the resulting list of property bag entries
            foreach (var property in entriesToDelete)
            {
                template.PropertyBagEntries.Remove(property);
            }

            return template;
        }

        private IEnumerable<T> SymmetricExcept<T>(IEnumerable<T> seq1, IEnumerable<T> seq2)
        {
            HashSet<T> hashSet = new HashSet<T>(seq1);
            hashSet.SymmetricExceptWith(seq2);
            return hashSet.Select(x => x);
        }


        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.PropertyBagEntries.Any() && !web.IsNoScriptSite();
            }
            return _willProvision.Value;

        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }
}
