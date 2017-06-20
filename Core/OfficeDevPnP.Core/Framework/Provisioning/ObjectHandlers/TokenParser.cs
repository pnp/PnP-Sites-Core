using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System.Resources;
using System.Collections;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Handles methods for token parser
    /// </summary>
    public class TokenParser
    {
        public Web _web;

        private List<TokenDefinition> _tokens;
        private List<Localization> _localizations = new List<Localization>();

        /// <summary>
        /// List of token definitions
        /// </summary>
        public List<TokenDefinition> Tokens
        {
            get { return _tokens; }
            private set
            {
                _tokens = value;
            }
        }

        /// <summary>
        /// adds token definition
        /// </summary>
        /// <param name="tokenDefinition">A TokenDefinition object</param>
        public void AddToken(TokenDefinition tokenDefinition)
        {

            _tokens.Add(tokenDefinition);
            // ORDER IS IMPORTANT!
            var sortedTokens = from t in _tokens
                               orderby t.GetTokenLength() descending
                               select t;

            _tokens = sortedTokens.ToList();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="web">A SharePoint site or subsite</param>
        /// <param name="template">a provisioning template</param>
        public TokenParser(Web web, ProvisioningTemplate template)
        {
            web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Language);

            _web = web;

            _tokens = new List<TokenDefinition>();

            _tokens.Add(new SiteCollectionToken(web));
            _tokens.Add(new SiteCollectionIdToken(web));
            _tokens.Add(new SiteToken(web));
            _tokens.Add(new MasterPageCatalogToken(web));
            _tokens.Add(new SiteCollectionTermStoreIdToken(web));
            _tokens.Add(new KeywordsTermStoreIdToken(web));
            _tokens.Add(new ThemeCatalogToken(web));
            _tokens.Add(new SiteNameToken(web));
            _tokens.Add(new SiteIdToken(web));
            _tokens.Add(new SiteOwnerToken(web));
            _tokens.Add(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.owners));
            _tokens.Add(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.members));
            _tokens.Add(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.visitors));
            _tokens.Add(new GuidToken(web));
            _tokens.Add(new DateNowToken(web));
            _tokens.Add(new CurrentUserIdToken(web));
            _tokens.Add(new CurrentUserLoginNameToken(web));
            _tokens.Add(new CurrentUserFullNameToken(web));
            _tokens.Add(new AuthenticationRealmToken(web));

            // Add lists
            web.Context.Load(web.Lists, ls => ls.Include(l => l.Id, l => l.Title, l => l.RootFolder.ServerRelativeUrl, l => l.Views));
            web.Context.ExecuteQueryRetry();
            foreach (var list in web.Lists)
            {
                _tokens.Add(new ListIdToken(web, list.Title, list.Id));
                _tokens.Add(new ListUrlToken(web, list.Title, list.RootFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length + 1)));

                foreach (var view in list.Views)
                {
                    _tokens.Add(new ListViewIdToken(web, list.Title, view.Title, view.Id));
                }
            }

            if (web.IsSubSite())
            {
                // Add lists from rootweb
                var rootWeb = (web.Context as ClientContext).Site.RootWeb;
                rootWeb.EnsureProperty(w => w.ServerRelativeUrl);
                rootWeb.Context.Load(rootWeb.Lists, ls => ls.Include(l => l.Id, l => l.Title, l => l.RootFolder.ServerRelativeUrl));
                rootWeb.Context.ExecuteQueryRetry();
                foreach (var rootList in rootWeb.Lists)
                {
                    // token already there? Skip the list
                    if (web.Lists.FirstOrDefault(l => l.Title == rootList.Title) == null)
                    {
                        _tokens.Add(new ListIdToken(web, rootList.Title, rootList.Id));
                        _tokens.Add(new ListUrlToken(web, rootList.Title, rootList.RootFolder.ServerRelativeUrl.Substring(rootWeb.ServerRelativeUrl.Length + 1)));
                    }
                }
            }

            // Add ContentTypes
            web.Context.Load(web.AvailableContentTypes, cs => cs.Include(ct => ct.StringId, ct => ct.Name));
            web.Context.ExecuteQueryRetry();
            foreach (var ct in web.AvailableContentTypes)
            {
                _tokens.Add(new ContentTypeIdToken(web, ct.Name, ct.StringId));
            }
            // Add parameters
            foreach (var parameter in template.Parameters)
            {
                _tokens.Add(new ParameterToken(web, parameter.Key, parameter.Value ?? string.Empty));
            }

            // Add TermSetIds
            TaxonomySession session = TaxonomySession.GetTaxonomySession(web.Context);

            var termStores = session.EnsureProperty(s => s.TermStores);
            foreach (var ts in termStores)
            {
                _tokens.Add(new TermStoreIdToken(web, ts.Name, ts.Id));
            }
            var termStore = session.GetDefaultSiteCollectionTermStore();
            web.Context.Load(termStore);
            web.Context.ExecuteQueryRetry();
            if (!termStore.ServerObjectIsNull.Value)
            {
                web.Context.Load(termStore.Groups,
                    g => g.Include(
                        tg => tg.Name,
                        tg => tg.TermSets.Include(
                            ts => ts.Name,
                            ts => ts.Id)
                    ));
                web.Context.ExecuteQueryRetry();
                foreach (var termGroup in termStore.Groups)
                {
                    foreach (var termSet in termGroup.TermSets)
                    {
                        _tokens.Add(new TermSetIdToken(web, termGroup.Name, termSet.Name, termSet.Id));
                    }
                }
            }

            _tokens.Add(new SiteCollectionTermGroupIdToken(web));
            _tokens.Add(new SiteCollectionTermGroupNameToken(web));

            // SiteCollection TermSets, only when we're not working in app-only
            if (!web.Context.IsAppOnly())
            {
                var site = (web.Context as ClientContext).Site;
                var siteCollectionTermGroup = termStore.GetSiteCollectionGroup(site, true);
                web.Context.Load(siteCollectionTermGroup);
                try
                {
                    web.Context.ExecuteQueryRetry();
                    if (null != siteCollectionTermGroup && !siteCollectionTermGroup.ServerObjectIsNull.Value)
                    {
                        web.Context.Load(siteCollectionTermGroup, group => group.TermSets.Include(ts => ts.Name, ts => ts.Id));
                        web.Context.ExecuteQueryRetry();
                        foreach (var termSet in siteCollectionTermGroup.TermSets)
                        {
                            _tokens.Add(new SiteCollectionTermSetIdToken(web, termSet.Name, termSet.Id));
                        }
                    }
                }
                catch (NullReferenceException)
                {
                    // If there isn't a default TermGroup for the Site Collection, we skip the terms in token handler
                }
            }

            // Fields
            var fields = web.Fields;
            web.Context.Load(fields, flds => flds.Include(f => f.Title, f => f.InternalName));
            web.Context.ExecuteQueryRetry();
            foreach (var field in fields)
            {
                _tokens.Add(new FieldTitleToken(web, field.InternalName, field.Title));
            }

            if (web.IsSubSite())
            {
                // SiteColumns from rootsite
                var rootWeb = (web.Context as ClientContext).Site.RootWeb;
                var siteColumns = rootWeb.Fields;
                web.Context.Load(siteColumns, flds => flds.Include(f => f.Title, f => f.InternalName));
                web.Context.ExecuteQueryRetry();
                foreach (var field in siteColumns)
                {
                    _tokens.Add(new FieldTitleToken(rootWeb, field.InternalName, field.Title));
                }
            }

            // Handle resources
            if (template.Localizations.Any())
            {
                // Read all resource keys in a list
                List<Tuple<string, uint, string>> resourceEntries = new List<Tuple<string, uint, string>>();
                foreach (var localizationEntry in template.Localizations)
                {
                    var filePath = localizationEntry.ResourceFile;
                    using (var stream = template.Connector.GetFileStream(filePath))
                    {
                        if (stream != null)
                        {
                            using (ResXResourceReader resxReader = new ResXResourceReader(stream))
                            {
                                foreach (DictionaryEntry entry in resxReader)
                                {
                                    resourceEntries.Add(new Tuple<string, uint, string>(entry.Key.ToString(), (uint)localizationEntry.LCID, entry.Value.ToString()));
                                }
                            }
                        }
                    }
                }

                var uniqueKeys = resourceEntries.Select(k => k.Item1).Distinct();
                foreach (var key in uniqueKeys)
                {
                    var matches = resourceEntries.Where(k => k.Item1 == key);
                    var entries = matches.Select(k => new ResourceEntry() { LCID = k.Item2, Value = k.Item3 }).ToList();
                    LocalizationToken token = new LocalizationToken(web, key, entries);

                    _tokens.Add(token);
                }
            }

            // OOTB Roledefs
            web.EnsureProperty(w => w.RoleDefinitions.Include(r => r.RoleTypeKind));
            foreach (var roleDef in web.RoleDefinitions.AsEnumerable().Where(r => r.RoleTypeKind != RoleType.None))
            {
                _tokens.Add(new RoleDefinitionToken(web, roleDef));
            }

            // Groups
            web.EnsureProperty(w => w.SiteGroups.Include(g => g.Title, g => g.Id));
            foreach (var siteGroup in web.SiteGroups)
            {
                _tokens.Add(new GroupIdToken(web, siteGroup.Title, siteGroup.Id));
            }
            web.EnsureProperty(w => w.AssociatedVisitorGroup).EnsureProperties(g => g.Id, g => g.Title);
            web.EnsureProperty(w => w.AssociatedMemberGroup).EnsureProperties(g => g.Id, g => g.Title);
            web.EnsureProperty(w => w.AssociatedOwnerGroup).EnsureProperties(g => g.Id, g => g.Title);

            if (!web.AssociatedVisitorGroup.ServerObjectIsNull.Value)
            {
                _tokens.Add(new GroupIdToken(web, "associatedvisitorgroup", web.AssociatedVisitorGroup.Id));
            }
            if (!web.AssociatedMemberGroup.ServerObjectIsNull.Value)
            {
                _tokens.Add(new GroupIdToken(web, "associatedmembergroup", web.AssociatedMemberGroup.Id));
            }
            if (!web.AssociatedOwnerGroup.ServerObjectIsNull.Value)
            {
                _tokens.Add(new GroupIdToken(web, "associatedownergroup", web.AssociatedOwnerGroup.Id));
            }
            var sortedTokens = from t in _tokens
                               orderby t.GetTokenLength() descending
                               select t;

            _tokens = sortedTokens.ToList();
        }

        /// <summary>
        /// Gets list of token resource values
        /// </summary>
        /// <param name="tokenValue">Token value</param>
        /// <returns>Returns list of token resource values</returns>
        public List<Tuple<string, string>> GetResourceTokenResourceValues(string tokenValue)
        {
            List<Tuple<string, string>> resourceValues = new List<Tuple<string, string>>();
            var resourceTokens = _tokens.Where(t => t is LocalizationToken && t.GetTokens().Contains(tokenValue));
            foreach (LocalizationToken resourceToken in resourceTokens)
            {
                var entries = resourceToken.ResourceEntries;
                foreach (var entry in entries)
                {
                    CultureInfo ci = new CultureInfo((int)entry.LCID);
                    resourceValues.Add(new Tuple<string, string>(ci.Name, entry.Value));
                }
            }
            return resourceValues;
        }

        /// <summary>
        /// Clears cache of tokens
        /// </summary>
        /// <param name="web">A SharePoint site or subsite</param>
        public void Rebase(Web web)
        {
            web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Language);

            _web = web;

            foreach (var token in _tokens)
            {
                token.ClearCache();
                token.Web = web;
            }
        }

        /// <summary>
        /// Parses the string
        /// </summary>
        /// <param name="input">input string to parse</param>
        /// <returns>Returns parsed string</returns>
        public string ParseString(string input)
        {
            return ParseString(input, null);
        }

        /// <summary>
        /// Gets left over tokens
        /// </summary>
        /// <param name="input">input string</param>
        /// <returns>Returns collections of left over tokens</returns>
        public IEnumerable<string> GetLeftOverTokens(string input)
        {
            List<string> values = new List<string>();
            var matches = Regex.Matches(input, "(?<guid>\\{\\S{8}-\\S{4}-\\S{4}-\\S{4}-\\S{12}?\\})").OfType<Match>().Select(m => m.Value);
            foreach (var match in matches)
            {
                Guid gout;
                if (!Guid.TryParse(match, out gout))
                {
                    values.Add(match);
                }
            }
            return values;
        }

        /// <summary>
        /// Parses given string
        /// </summary>
        /// <param name="input">input string</param>
        /// <param name="tokensToSkip">array of tokens to skip</param>
        /// <returns>Returns parsed string</returns>
        public string ParseString(string input, params string[] tokensToSkip)
        {
            var origInput = input;
            if (!string.IsNullOrEmpty(input))
            {
                foreach (var token in _tokens)
                {
                    if (tokensToSkip != null)
                    {
                        var filteredTokens = token.GetTokens().Except(tokensToSkip, StringComparer.InvariantCultureIgnoreCase);
                        if (filteredTokens.Any())
                        {
                            foreach (var filteredToken in filteredTokens)
                            {
                                var regex = token.GetRegexForToken(filteredToken);
                                if (regex.IsMatch(input))
                                {
                                    input = regex.Replace(input, token.GetReplaceValue());
                                }
                            }
                        }
                    }
                    else
                    {
                        var matchingTokens = token.GetRegex().Where(regex => regex.IsMatch(input));
                        foreach (var regex in matchingTokens)
                        {
                            input = regex.Replace(input, token.GetReplaceValue());
                        }
                    }
                }
            }

            while (origInput != input)
            {
                foreach (var token in _tokens)
                {
                    origInput = input;
                    if (tokensToSkip != null)
                    {
                        var filteredTokens = token.GetTokens().Except(tokensToSkip, StringComparer.InvariantCultureIgnoreCase);
                        if (filteredTokens.Any())
                        {
                            foreach (var filteredToken in filteredTokens)
                            {
                                var regex = token.GetRegexForToken(filteredToken);
                                if (regex.IsMatch(input))
                                {
                                    input = regex.Replace(input, token.GetReplaceValue());
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (var regex in token.GetRegex().Where(regex => regex.IsMatch(input)))
                        {
                            origInput = input;
                            input = regex.Replace(input, token.GetReplaceValue());
                        }
                    }
                }
            }

            return input;
        }
    }
}

