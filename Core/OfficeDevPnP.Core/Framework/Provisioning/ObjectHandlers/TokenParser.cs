using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class TokenParser
    {
        public Web _web;

        private List<TokenDefinition> _tokens = new List<TokenDefinition>();

        public List<TokenDefinition> Tokens
        {
            get { return _tokens; }
            private set
            {
                _tokens = value;
            }
        }

        public void AddToken(TokenDefinition tokenDefinition)
        {

            _tokens.Add(tokenDefinition);
            // ORDER IS IMPORTANT!
            var sortedTokens = from t in _tokens
                               orderby t.GetTokenLength() descending
                               select t;

            _tokens = sortedTokens.ToList();
        }

        public TokenParser(Web web, ProvisioningTemplate template)
        {
            web.EnsureProperties(w => w.ServerRelativeUrl);

            _web = web;

            _tokens = new List<TokenDefinition>();

            _tokens.Add(new SiteCollectionToken(web));
            _tokens.Add(new SiteToken(web));
            _tokens.Add(new MasterPageCatalogToken(web));
            _tokens.Add(new SiteCollectionTermStoreIdToken(web));
            _tokens.Add(new KeywordsTermStoreIdToken(web));
            _tokens.Add(new ThemeCatalogToken(web));
            _tokens.Add(new SiteNameToken(web));
            _tokens.Add(new SiteIdToken(web));
            _tokens.Add(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.owners));
            _tokens.Add(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.members));
            _tokens.Add(new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.visitors));
            _tokens.Add(new GuidToken(web));
            _tokens.Add(new CurrentUserIdToken(web));
            _tokens.Add(new CurrentUserLoginNameToken(web));
            _tokens.Add(new CurrentUserFullNameToken(web));

            // Add lists
            web.Context.Load(web.Lists, ls => ls.Include(l => l.Id, l => l.Title, l => l.RootFolder.ServerRelativeUrl));
            web.Context.ExecuteQueryRetry();
            foreach (var list in web.Lists)
            {
                _tokens.Add(new ListIdToken(web, list.Title, list.Id));
                _tokens.Add(new ListUrlToken(web, list.Title, list.RootFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length + 1)));
            }

            // Add ContentTypes
            web.Context.Load(web.ContentTypes, cs => cs.Include(ct => ct.StringId, ct => ct.Name));
            web.Context.ExecuteQueryRetry();
            foreach (var ct in web.ContentTypes)
            {
                _tokens.Add(new ContentTypeIdToken(web, ct.Name, ct.StringId));
            }
            // Add parameters
            foreach (var parameter in template.Parameters)
            {
                _tokens.Add(new ParameterToken(web, parameter.Key, parameter.Value));
            }

            // Add TermSetIds
            TaxonomySession session = TaxonomySession.GetTaxonomySession(web.Context);

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

            var sortedTokens = from t in _tokens
                               orderby t.GetTokenLength() descending
                               select t;

            _tokens = sortedTokens.ToList();
        }

        public void Rebase(Web web)
        {
            _web = web;

            foreach (var token in _tokens)
            {
                token.ClearCache();
                token.Web = web;
            }
        }

        public string ParseString(string input)
        {
            return ParseString(input, null);
        }

        public string ParseString(string input, params string[] tokensToSkip)
        {
            var origInput = input;
            if (!string.IsNullOrEmpty(input))
            {
                foreach (var token in _tokens)
                {
                    if (tokensToSkip != null)
                    {
                        if (token.GetTokens().Except(tokensToSkip, StringComparer.InvariantCultureIgnoreCase).Any())
                        {
                            foreach (var filteredToken in token.GetTokens().Except(tokensToSkip, StringComparer.InvariantCultureIgnoreCase))
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
                        if (token.GetTokens().Except(tokensToSkip, StringComparer.InvariantCultureIgnoreCase).Any())
                        {
                            foreach (var filteredToken in token.GetTokens().Except(tokensToSkip, StringComparer.InvariantCultureIgnoreCase))
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
