using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
#if !SP2013
    class LocalizationValidator : ValidatorBase
    {
        private bool isNoScriptSite = false;

        #region construction
        public LocalizationValidator(Web web) : base()
        {
            // optionally override schema version
            // SchemaVersion = "http://schemas.dev.office.com/PnP/2016/05/ProvisioningSchema";
            // XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:ContentTypes/pnp:ContentType";

            isNoScriptSite = web.IsNoScriptSite();
            cc = web.Context as ClientContext;
        }
        #endregion

        public bool Validate(ProvisioningTemplate ptSource, ProvisioningTemplate ptTarget, TokenParser sParser, TokenParser tParser, Web web)
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

            #region ListViews
            if (CanUseAcceptLanguageHeaderForLocalization(web))
            {
                isValid = ValidateListView(ptSource, sParser);
                if (!isValid) return false;
            }
            #endregion

            #region WebParts
            if (!isNoScriptSite && CanUseAcceptLanguageHeaderForLocalization(web))
            {
                isValid = ValidateWebPartOnPages(ptSource, sParser);
                if (!isValid) { return false; }
            }
            #endregion

            #region Navigation
            if (!isNoScriptSite && CanUseAcceptLanguageHeaderForLocalization(web))
            {
                isValid = ValidateStructuralNavigation(ptSource, sParser);
                if (!isValid) return false;
            }
            #endregion

#if !SP2013 && !SP2016
            #region Custom Action
            if (!isNoScriptSite)
            {
                isValid = ValidateCustomActions(ptSource.CustomActions, ptTarget.CustomActions, sParser, tParser, ptTarget.SiteFields.Count > 0);
                if (!isValid) { return false; }
            }
            #endregion
#endif
            return isValid;
        }

        #region WebParts
        private bool CanUseAcceptLanguageHeaderForLocalization(Web web)
        {
            if (web.Context.IsAppOnly())
            {
                return true;
            }

            var currentUser = web.EnsureProperty(w => w.CurrentUser);
            PeopleManager peopleManager = new PeopleManager(web.Context);
            var languageSettings = peopleManager.GetUserProfilePropertyFor(web.CurrentUser.LoginName, "SPS-MUILanguages");
            web.Context.ExecuteQueryRetry();

            if (languageSettings == null || String.IsNullOrEmpty(languageSettings.Value))
            {
                return true;
            }

            return false;
        }

        public bool ValidateWebPartOnPages(ProvisioningTemplate template, TokenParser parser)
        {
            var web = cc.Web;
            var file = template.Files.First();
            var folderName = parser.ParseString(file.Folder);
            var url = folderName + "/" + template.Connector.GetFilenamePart(file.Src);
            var resourceValues = parser.GetResourceTokenResourceValues(file.WebParts.First().Title);
            var ok = ValidatePartOnPage(parser, resourceValues, web, url);
            if (!ok) return false;

            var page = template.Pages.First();
            url = parser.ParseString(page.Url);
            resourceValues = parser.GetResourceTokenResourceValues(file.WebParts.First().Title);
            ok = ValidatePartOnPage(parser, resourceValues, web, url);
            return ok;
        }

        private bool ValidatePartOnPage(TokenParser parser, IEnumerable<Tuple<string, string>> resourceValues, Web web, string url)
        {
            bool allOk = true;
            foreach (var resourceValue in resourceValues)
            {
                var webPartDef = web.GetWebParts(parser.ParseString(url)).First();
                webPartDef.Context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                webPartDef.EnsureProperties(p => p.WebPart, p => p.WebPart.Properties);
                if (!webPartDef.WebPart.Properties["Title"].Equals(resourceValue.Item2))
                {
                    allOk = false;
                }
            }
            web.Context.PendingRequest.RequestExecutor.WebRequest.Headers.Remove("Accept-Language");
            return allOk;
        }
        #endregion

        #region SiteFields
        public bool ValidateSiteFields(Core.Framework.Provisioning.Model.FieldCollection sElements, Core.Framework.Provisioning.Model.FieldCollection tElements, TokenParser sParser, TokenParser tParser)
        {
            List<Localization> sColl = LoadFields(sElements);
            List<Localization> tColl = LoadFields(tElements);

            if (sColl.Count > 0)
            {
                if (!Validatelocalization(sColl, tColl, sParser, tParser)) { return false; }
            }

            return true;
        }

        private List<Localization> LoadFields(Core.Framework.Provisioning.Model.FieldCollection coll)
        {
            string attribute1 = "DisplayName";
            string attribute2 = "Description";
            string key = "ID";
            List<Localization> loc = new List<Localization>();

            foreach (Core.Framework.Provisioning.Model.Field item in coll)
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

        #region ListViews
        public bool ValidateListView(ProvisioningTemplate template, TokenParser parser)
        {
            var web = cc.Web;
            cc.Load(web.Lists);
            cc.ExecuteQueryRetry();
            var allOk = true;
            foreach (var listDef in template.Lists)
            {
                var list = web.GetListByUrl(listDef.Url);
                foreach (var viewDef in listDef.Views)
                {
                    XElement currentXml = XElement.Parse(viewDef.SchemaXml);
                    var viewUrl = currentXml.Attribute("Url").Value;
                    var dispName = currentXml.Attribute("DisplayName").Value;
                    if (dispName.ContainsResourceToken())
                    {
                        var resourceValues = parser.GetResourceTokenResourceValues(dispName);
                        foreach (var resourceValue in resourceValues)
                        {
                            list.Context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                            list.Context.Load(list.Views);
                            list.Context.ExecuteQueryRetry();
                            var view = list.Views.Single(v => v.ServerRelativeUrl.EndsWith(viewUrl));
                            if (!view.Title.Equals(resourceValue.Item2))
                            {
                                allOk = false;
                            }
                        }
                    }
                }

            }
            return allOk;
        }
        #endregion

        #region ContenType Tests
        public bool ValidateContentTypes(Core.Framework.Provisioning.Model.ContentTypeCollection sElements, Core.Framework.Provisioning.Model.ContentTypeCollection tElements, TokenParser sParser, TokenParser tParser)
        {
            List<Localization> sColl = LoadContentTypes(sElements);
            List<Localization> tColl = LoadContentTypes(tElements);

            if (sColl.Count > 0)
            {
                if (!Validatelocalization(sColl, tColl, sParser, tParser)) { return false; }
            }

            return true;
        }

        private List<Localization> LoadContentTypes(Core.Framework.Provisioning.Model.ContentTypeCollection coll)
        {
            List<Localization> loc = new List<Localization>();

            foreach (Core.Framework.Provisioning.Model.ContentType item in coll)
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

        #region Navigation
        public bool ValidateStructuralNavigation(ProvisioningTemplate template, TokenParser parser)
        {
            bool ok = true;
            var web = cc.Web;
            if (template.Navigation == null) return true;
            if (template.Navigation.GlobalNavigation == null) return true;
            if (template.Navigation.GlobalNavigation.NavigationType == GlobalNavigationType.Managed) return true;

            var node = template.Navigation.GlobalNavigation.StructuralNavigation.NavigationNodes.First();
            if (node.Title.ContainsResourceToken())
            {
                var resourceValues = parser.GetResourceTokenResourceValues(node.Title);
                foreach (var resourceValue in resourceValues)
                {
                    cc.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                    cc.Load(web, w => w.Navigation, w => w.Navigation.TopNavigationBar);
                    cc.ExecuteQueryRetry();
                    if (!cc.Web.IsSubSite())
                    {
                        var firstNode = web.Navigation.TopNavigationBar.First();
                        if (!firstNode.Title.Equals(resourceValue.Item2))
                        {
                            ok = false;
                        }
                    }
                }
            }
            return ok;
        }
        #endregion

        public class Localization
        {
            public Localization(string key, string title, string description)
            {
                Key = key;
                Title = title;
                Description = description;
            }

            public string Key { get; set; }

            public string Title { get; set; }

            public string Description { get; set; }

            public Core.Framework.Provisioning.Model.FieldCollection Fields { get; set; }
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
