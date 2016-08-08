using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
#if !SP2013
    class LocalizationValidator : ValidatorBase
    {
        #region construction
        public LocalizationValidator() : base()
        {
            // optionally override schema version
            // SchemaVersion = "http://schemas.dev.office.com/PnP/2016/05/ProvisioningSchema";
            // XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:ContentTypes/pnp:ContentType";
        }
        #endregion

        public bool Validate(ProvisioningTemplate ptSource, ProvisioningTemplate ptTarget, TokenParser sParser, TokenParser tParser)
        {
            bool isValid = false;

            #region SiteFields
            if (ptTarget.SiteFields.Count > 0)
            {
                isValid = ValidateSiteFields(ptSource.SiteFields, ptTarget.SiteFields, sParser, tParser);
                if (!isValid) { return false; }
            }
            #endregion

            #region ContentTypes
            if (ptTarget.ContentTypes.Count > 0)
            {
                isValid = ValidateContentTypes(ptSource.ContentTypes, ptTarget.ContentTypes, sParser, tParser);
                if (!isValid) { return false; }
            }
            #endregion

            #region ListInstances
            isValid = ValidateListInstances(ptSource.Lists, ptTarget.Lists, sParser, tParser);
            if (!isValid) { return false; }
            #endregion

#if !ONPREMISES
            #region Custom Action
            isValid = ValidateCustomActions(ptSource.CustomActions, ptTarget.CustomActions, sParser, tParser, ptTarget.SiteFields.Count > 0);
            if (!isValid) { return false; }
            #endregion
#endif
            return isValid;
        }
    
        #region SiteFields
        public bool ValidateSiteFields(FieldCollection sElements, FieldCollection tElements, TokenParser sParser, TokenParser tParser)
        {
            List<Localization> sColl = LoadFields(sElements);
            List<Localization> tColl = LoadFields(tElements);

            if (sColl.Count > 0)
            {
                if (!Validatelocalization(sColl, tColl, sParser, tParser)) { return false; }
            }

            return true;
        }

        private List<Localization> LoadFields(FieldCollection coll)
        {
            string attribute1 = "DisplayName";
            string attribute2 = "Description";
            string key = "ID";
            List<Localization> loc = new List<Localization>();

            foreach (Field item in coll)
            {
                XElement element = XElement.Parse(item.SchemaXml);
                string sTokenValue1 = GetPropertyValue(attribute1, element);
                string sTokenValue2 = GetPropertyValue(attribute2, element);
                string sKey = GetPropertyValue(key, element);

                loc.Add(new Localization(sKey, sTokenValue1, sTokenValue2));
            }

            return loc;
        }

        private string GetPropertyValue(string attribute, XElement element)
        {
            return element.Attribute(attribute) != null ? element.Attribute(attribute).Value : "";
        }
        #endregion

        #region ContenType Tests
        public bool ValidateContentTypes(ContentTypeCollection sElements, ContentTypeCollection tElements, TokenParser sParser, TokenParser tParser)
        {
            List<Localization> sColl = LoadContentTypes(sElements);
            List<Localization> tColl = LoadContentTypes(tElements);

            if (sColl.Count > 0)
            {
                if (!Validatelocalization(sColl, tColl, sParser, tParser)) { return false; }
            }

            return true;
        }

        private List<Localization> LoadContentTypes(ContentTypeCollection coll)
        {
            List<Localization> loc = new List<Localization>();

            foreach (ContentType item in coll)
            {
                loc.Add(new Localization(item.Id, item.Name, item.Description));
            }           

            return loc;
        }
        #endregion

        #region ListInstances        
        public bool ValidateListInstances(ListInstanceCollection sElements, ListInstanceCollection tElements, TokenParser sParser, TokenParser tParser)
        {
            List<Localization> sColl = LoadListInstances(sElements);
            List<Localization> tColl = LoadListInstances(tElements);

            if (sColl.Count > 0)
            {
                if (!Validatelocalization(sColl, tColl, sParser, tParser)) { return false; }
            }

            return true;
        }

        private List<Localization> LoadListInstances(ListInstanceCollection coll)
        {
            List<Localization> loc = new List<Localization>();

            foreach (ListInstance item in coll)
            {
                Localization localization = new Localization(item.Url, item.Title, item.Description);
                if (item.Fields != null) { localization.Fields = item.Fields; }
                loc.Add(localization);
            }

            return loc;
        }
        #endregion

        #region CustomActions
        private bool ValidateCustomActions(CustomActions srcCustomActions, CustomActions targetCustomActions, TokenParser sParser, TokenParser tParser, bool rootSite)
        {
            List<Localization> sCustomActions = LoadCustomActions(srcCustomActions, rootSite);
            List<Localization> tCustomActions = LoadCustomActions(targetCustomActions, rootSite);

            if (sCustomActions.Count > 0)
            {
                if (!Validatelocalization(sCustomActions, tCustomActions, sParser, tParser)) { return false; }
            }

            return true;
        }

        private List<Localization> LoadCustomActions(CustomActions customActions, bool rootSite)
        {
            List<Localization> locCustomActions = new List<Localization>();

            if (rootSite)
            {
                foreach (CustomAction action in customActions.SiteCustomActions)
                {
                    locCustomActions.Add(new Localization(action.Name, action.Title, action.Description));
                }
            }
            foreach (CustomAction action in customActions.WebCustomActions)
            {
                locCustomActions.Add(new Localization(action.Name, action.Title, action.Description));
            }

            return locCustomActions;
        }     
        #endregion

        public class Localization {
            public Localization(string key, string title, string description)
            {
                Key = key;
                Title = title;
                Description = description;
            }

            public string Key { get; set; }

            public string Title { get; set; }

            public string Description { get; set; }

            public FieldCollection Fields { get; set; }
        }

        private bool Validatelocalization(List<Localization> sElements, List<Localization> tElements, TokenParser sParser, TokenParser tParser)
        {
            bool isValid = false;
            int sCount = 0;
            int tCount = 0;

            foreach (var sElement in sElements)
            {
                var sTokenValue1 = sElement.Title;
                var sTokenValue2 = sElement.Description;
                string sKey = sElement.Key;

                bool isProp1ContainsRes = sTokenValue1.ContainsResourceToken();
                bool isProp2ContainsRes = sTokenValue2.ContainsResourceToken();

                if (isProp1ContainsRes || isProp2ContainsRes)
                {
                    sCount++;
                    foreach (var tElement in tElements)
                    {
                        var tKey = tElement.Key;
                        if (sKey.ToLower() == tKey.ToLower())
                        {
                            if (isProp1ContainsRes)
                            {
                                if (!ValidateResourceEntries(sParser, tParser, sTokenValue1, tElement.Title)) { return false; }
                            }

                            if (isProp2ContainsRes)
                            {
                                if (!ValidateResourceEntries(sParser, tParser, sTokenValue2, tElement.Description)) { return false; }
                            }

                            if (sElement.Fields != null) // validate if list contains fields
                            {
                                if (!ValidateSiteFields(sElement.Fields, tElement.Fields, sParser, tParser)) { return false; }
                            }

                            tCount++;
                            break;
                        }
                    }
                }
            }

            if (sCount == tCount) { isValid = true; }
            return isValid;
        }

        private bool ValidateResourceEntries(TokenParser sParser, TokenParser tParser, string sTokenValue, string tTokenValue)
        {
            int sCount = 0;
            int tCount = 0;
            bool isValid = false;

            var sResValues = sParser.GetResourceTokenResourceValues(sTokenValue);
            var tResValues = tParser.GetResourceTokenResourceValues(tTokenValue);

            foreach (var sResVal in sResValues)
            {
                sCount++;
                var tResItem = tResValues.Where(t => t.Item1.ToLower().Equals(sResVal.Item1.ToLower())).FirstOrDefault();
                if (tResItem != null && tResItem.Item2.ToLower() == sResVal.Item2.ToLower())
                {
                    tCount++;
                    break;
                }
            }

            if (sCount == tCount) { isValid = true; }
            return isValid;
        }
    }
#endif
}
