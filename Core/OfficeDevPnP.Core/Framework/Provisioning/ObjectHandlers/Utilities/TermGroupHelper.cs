using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using static OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.ObjectTermGroups;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    internal static class TermGroupHelper
    {
        internal static List<ReusedTerm> ProcessGroup(Web web, TermStore termStore, Model.TermGroup modelTermGroup, TermGroup siteCollectionTermGroup, TokenParser parser, PnPMonitoredScope scope)
        {
            List<ReusedTerm> reusedTerms = new List<ReusedTerm>();

            SiteCollectionTermGroupNameToken siteCollectionTermGroupNameToken =
              new SiteCollectionTermGroupNameToken(web);

            #region Group

            var newGroup = false;
            var normalizedGroupName = TaxonomyItem.NormalizeName(web.Context, modelTermGroup.Name);
            web.Context.ExecuteQueryRetry();

            TermGroup group = termStore.Groups.FirstOrDefault(
                g => g.Id == modelTermGroup.Id || g.Name == normalizedGroupName.Value);
            if (group == null)
            {
                var parsedGroupName = parser.ParseString(modelTermGroup.Name);
                var parsedDescription = parser.ParseString(modelTermGroup.Description);

                if (modelTermGroup.Name == "Site Collection" ||
                    parsedGroupName == siteCollectionTermGroupNameToken.GetReplaceValue() ||
                    modelTermGroup.SiteCollectionTermGroup)
                {
                    var site = (web.Context as ClientContext).Site;
                    group = termStore.GetSiteCollectionGroup(site, true);
                    web.Context.Load(group, g => g.Name, g => g.Id, g => g.TermSets.Include(
                        tset => tset.Name,
                        tset => tset.Id));
                    web.Context.ExecuteQueryRetry();
                }
                else
                {
                    var parsedNormalizedGroupName = TaxonomyItem.NormalizeName(web.Context, parsedGroupName);
                    web.Context.ExecuteQueryRetry();

                    group = termStore.Groups.FirstOrDefault(g => g.Name == parsedNormalizedGroupName.Value);

                    if (group == null)
                    {
                        if (modelTermGroup.Id == Guid.Empty)
                        {
                            modelTermGroup.Id = Guid.NewGuid();
                        }
                        group = termStore.CreateGroup(parsedGroupName, modelTermGroup.Id);

                        group.Description = parsedDescription;

#if !ONPREMISES

                        // Handle TermGroup Contributors, if any
                        if (modelTermGroup.Contributors != null && modelTermGroup.Contributors.Count > 0)
                        {
                            foreach (var c in modelTermGroup.Contributors)
                            {
                                group.AddContributor(c.Name);
                            }
                        }

                        // Handle TermGroup Managers, if any
                        if (modelTermGroup.Managers != null && modelTermGroup.Managers.Count > 0)
                        {
                            foreach (var m in modelTermGroup.Managers)
                            {
                                group.AddGroupManager(m.Name);
                            }
                        }

#endif

                        termStore.CommitAll();
                        web.Context.Load(group);
                        web.Context.ExecuteQueryRetry();

                        newGroup = true;

                    }
                }

            }
            #endregion

            #region TermSets

            foreach (var modelTermSet in modelTermGroup.TermSets)
            {
                TermSet set = null;
                var newTermSet = false;

                var normalizedTermSetName = TaxonomyItem.NormalizeName(web.Context, modelTermSet.Name);
                web.Context.ExecuteQueryRetry();

                if (!newGroup)
                {
                    set =
                        group.TermSets.FirstOrDefault(
                            ts => ts.Id == modelTermSet.Id || ts.Name == normalizedTermSetName.Value);
                }
                if (set == null)
                {
                    if (modelTermSet.Id == Guid.Empty)
                    {
                        modelTermSet.Id = Guid.NewGuid();
                    }
                    set = group.CreateTermSet(parser.ParseString(modelTermSet.Name), modelTermSet.Id,
                        modelTermSet.Language ?? termStore.DefaultLanguage);
                    parser.AddToken(new TermSetIdToken(web, group.Name, modelTermSet.Name, modelTermSet.Id));
                    if (!siteCollectionTermGroup.ServerObjectIsNull.Value)
                    {
                        if (group.Name == siteCollectionTermGroup.Name)
                        {
                            parser.AddToken((new SiteCollectionTermSetIdToken(web, modelTermSet.Name, modelTermSet.Id)));
                        }
                    }
                    newTermSet = true;
                    set.Description = parser.ParseString(modelTermSet.Description);
                    set.IsOpenForTermCreation = modelTermSet.IsOpenForTermCreation;
                    set.IsAvailableForTagging = modelTermSet.IsAvailableForTagging;
                    foreach (var property in modelTermSet.Properties)
                    {
                        set.SetCustomProperty(property.Key, parser.ParseString(property.Value));
                    }
                    if (modelTermSet.Owner != null)
                    {
                        set.Owner = modelTermSet.Owner;
                    }
                    termStore.CommitAll();
                    web.Context.Load(set);
                    web.Context.ExecuteQueryRetry();
                }

                web.Context.Load(set, s => s.Terms.Include(t => t.Id, t => t.Name));
                web.Context.ExecuteQueryRetry();
                var terms = set.Terms;

                foreach (var modelTerm in modelTermSet.Terms)
                {
                    if (!newTermSet)
                    {
                        if (terms.Any())
                        {
                            var term = terms.FirstOrDefault(t => t.Id == modelTerm.Id);
                            if (term == null)
                            {
                                var normalizedTermName = TaxonomyItem.NormalizeName(web.Context, modelTerm.Name);
                                web.Context.ExecuteQueryRetry();

                                term = terms.FirstOrDefault(t => t.Name == normalizedTermName.Value);
                                if (term == null)
                                {
                                    var returnTuple = CreateTerm<TermSet>(web, modelTerm, set, termStore, parser, scope);
                                    if (returnTuple != null)
                                    {
                                        modelTerm.Id = returnTuple.Item1;
                                        parser = returnTuple.Item2;
                                    }
                                }
                                else
                                {
                                    modelTerm.Id = term.Id;
                                }
                            }
                            else
                            {
                                modelTerm.Id = term.Id;
                            }

                            if (term != null)
                            {
                                CheckChildTerms(web, modelTerm, term, termStore, parser, scope);
                            }
                        }
                        else
                        {
                            var returnTuple = CreateTerm<TermSet>(web, modelTerm, set, termStore, parser, scope);
                            if (returnTuple != null)
                            {
                                modelTerm.Id = returnTuple.Item1;
                                parser = returnTuple.Item2;
                            }
                        }
                    }
                    else
                    {
                        var returnTuple = CreateTerm<TermSet>(web, modelTerm, set, termStore, parser, scope);
                        if (returnTuple != null)
                        {
                            modelTerm.Id = returnTuple.Item1;
                            parser = returnTuple.Item2;
                        }
                    }
                }

                // do we need custom sorting?
                if (modelTermSet.Terms.Any(t => t.CustomSortOrder > -1))
                {
                    var sortedTerms = modelTermSet.Terms.OrderBy(t => t.CustomSortOrder);

                    var customSortString = sortedTerms.Aggregate(string.Empty,
                        (a, i) => a + i.Id.ToString() + ":");
                    customSortString = customSortString.TrimEnd(new[] { ':' });

                    set.CustomSortOrder = customSortString;
                    termStore.CommitAll();
                    web.Context.ExecuteQueryRetry();
                }
            }

            #endregion

            return reusedTerms;
        }

        internal class ReusedTerm
        {
            public Model.Term ModelTerm { get; set; }
            public TaxonomyItem Parent { get; set; }
            public TermStore TermStore { get; set; }
        }

        internal static Tuple<Guid, TokenParser, List<ReusedTerm>> CreateTerm<T>(Web web, Model.Term modelTerm, TaxonomyItem parent,
           TermStore termStore, TokenParser parser, PnPMonitoredScope scope) where T : TaxonomyItem
        {

            var reusedTerms = new List<ReusedTerm>();
            // If the term is a re-used term and the term is not a source term, skip for now and create later
            if (modelTerm.IsReused && !modelTerm.IsSourceTerm)
            {
                reusedTerms.Add(new ReusedTerm()
                {
                    ModelTerm = modelTerm,
                    Parent = parent,
                    TermStore = termStore
                });
                return Tuple.Create(Guid.Empty, parser, reusedTerms); ;
            }

            // Create new term
            Term term;
            if (modelTerm.Id == Guid.Empty)
            {
                modelTerm.Id = Guid.NewGuid();
            }

            if (parent is Term)
            {
                term = ((Term)parent).CreateTerm(parser.ParseString(modelTerm.Name), modelTerm.Language != null && modelTerm.Language != 0 ? modelTerm.Language.Value : termStore.DefaultLanguage, modelTerm.Id);
            }
            else
            {
                term = ((TermSet)parent).CreateTerm(parser.ParseString(modelTerm.Name), modelTerm.Language != null && modelTerm.Language != 0 ? modelTerm.Language.Value : termStore.DefaultLanguage, modelTerm.Id);
            }
            if (!string.IsNullOrEmpty(modelTerm.Description))
            {
                term.SetDescription(parser.ParseString(modelTerm.Description), modelTerm.Language != null && modelTerm.Language != 0 ? modelTerm.Language.Value : termStore.DefaultLanguage);
            }
            if (!string.IsNullOrEmpty(modelTerm.Owner))
            {
                term.Owner = modelTerm.Owner;
            }

            term.IsAvailableForTagging = modelTerm.IsAvailableForTagging;

            if (modelTerm.Properties.Any() || modelTerm.Labels.Any() || modelTerm.LocalProperties.Any())
            {
                if (modelTerm.Labels.Any())
                {
                    CreateTermLabels(modelTerm, termStore, parser, scope, term);
                }

                if (modelTerm.Properties.Any())
                {
                    SetTermCustomProperties(modelTerm, parser, term);
                }
                if (modelTerm.LocalProperties.Any())
                {
                    SetTermLocalCustomProperties(modelTerm, parser, term);
                }
            }

            termStore.CommitAll();

            web.Context.Load(term);
            web.Context.ExecuteQueryRetry();

            // Deprecate term if needed
            if (modelTerm.IsDeprecated != term.IsDeprecated)
            {
                term.Deprecate(modelTerm.IsDeprecated);
                web.Context.ExecuteQueryRetry();
            }


            parser = CreateChildTerms(web, modelTerm, term, termStore, parser, scope);
            return Tuple.Create(modelTerm.Id, parser, reusedTerms);
        }


        private static void CreateTermLabels(Model.Term modelTerm, TermStore termStore, TokenParser parser, PnPMonitoredScope scope, Term term)
        {
            foreach (var label in modelTerm.Labels)
            {
                if (((label.IsDefaultForLanguage && label.Language != termStore.DefaultLanguage) || label.IsDefaultForLanguage == false) && termStore.Languages.Contains(label.Language))
                {
                    term.CreateLabel(parser.ParseString(label.Value), label.Language, label.IsDefaultForLanguage);
                }
                else
                {
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language);
                }
            }
        }

        private static void SetTermCustomProperties(Model.Term modelTerm, TokenParser parser, Term term)
        {
            foreach (var property in modelTerm.Properties)
            {
                term.SetCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
            }
        }

        private static void SetTermLocalCustomProperties(Model.Term modelTerm, TokenParser parser, Term term)
        {
            foreach (var property in modelTerm.LocalProperties)
            {
                term.SetLocalCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
            }
        }

        /// <summary>
        /// Creates child terms for the current model term if any exist
        /// </summary>
        /// <param name="web"></param>
        /// <param name="modelTerm"></param>
        /// <param name="term"></param>
        /// <param name="termStore"></param>
        /// <param name="parser"></param>
        /// <param name="scope"></param>
        /// <returns>Updated parser object</returns>
        private static TokenParser CreateChildTerms(Web web, Model.Term modelTerm, Term term, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
        {
            if (modelTerm.Terms.Any())
            {
                foreach (var modelTermTerm in modelTerm.Terms)
                {
                    web.Context.Load(term.Terms);
                    web.Context.ExecuteQueryRetry();
                    var termTerms = term.Terms;
                    if (termTerms.Any())
                    {
                        var termTerm = termTerms.FirstOrDefault(t => t.Id == modelTermTerm.Id);
                        if (termTerm == null)
                        {
                            termTerm = termTerms.FirstOrDefault(t => t.Name == modelTermTerm.Name);
                            if (termTerm == null)
                            {
                                var returnTuple = CreateTerm<Term>(web, modelTermTerm, term, termStore, parser, scope);
                                if (returnTuple != null)
                                {
                                    modelTermTerm.Id = returnTuple.Item1;
                                    parser = returnTuple.Item2;
                                }
                            }
                            else
                            {
                                modelTermTerm.Id = termTerm.Id;
                            }
                        }
                        else
                        {
                            modelTermTerm.Id = termTerm.Id;
                        }
                    }
                    else
                    {
                        var returnTuple = CreateTerm<Term>(web, modelTermTerm, term, termStore, parser, scope);
                        if (returnTuple != null)
                        {
                            modelTermTerm.Id = returnTuple.Item1;
                            parser = returnTuple.Item2;
                        }
                    }
                }
                if (modelTerm.Terms.Any(t => t.CustomSortOrder > -1))
                {
                    var sortedTerms = modelTerm.Terms.OrderBy(t => t.CustomSortOrder);

                    var customSortString = sortedTerms.Aggregate(string.Empty, (a, i) => a + i.Id.ToString() + ":");
                    customSortString = customSortString.TrimEnd(new[] { ':' });

                    term.CustomSortOrder = customSortString;
                    termStore.CommitAll();
                }
            }

            return parser;
        }

        /// <summary>
        /// Attempts to reuse the model term. If the term does not yet exists it will return
        /// false for the first part of the the return tuple. this will notify the system
        /// that the term should be created instead of re-used.
        /// </summary>
        /// <param name="web"></param>
        /// <param name="modelTerm"></param>
        /// <param name="parent"></param>
        /// <param name="termStore"></param>
        /// <param name="parser"></param>
        /// <param name="scope"></param>
        /// <returns></returns>
        internal static TryReuseTermResult TryReuseTerm(Web web, Model.Term modelTerm, TaxonomyItem parent, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
        {
            if (!modelTerm.IsReused) return new TryReuseTermResult() { Success = false, UpdatedParser = parser };
            if (modelTerm.Id == Guid.Empty) return new TryReuseTermResult() { Success = false, UpdatedParser = parser };

            // Since we're reusing terms ensure the previous terms are committed
            termStore.CommitAll();
            web.Context.ExecuteQueryRetry();

            // Try to retrieve a matching term from the website also marked from re-use.  
            var taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
            web.Context.Load(taxonomySession);
            web.Context.ExecuteQueryRetry();

            if (taxonomySession.ServerObjectIsNull())
            {
                return new TryReuseTermResult() { Success = false, UpdatedParser = parser };
            }

            var freshTermStore = taxonomySession.GetDefaultKeywordsTermStore();
            Term preExistingTerm = freshTermStore.GetTerm(modelTerm.Id);

            try
            {
                web.Context.Load(preExistingTerm);
                web.Context.ExecuteQueryRetry();

                if (preExistingTerm.ServerObjectIsNull())
                {
                    preExistingTerm = null;
                }
            }
            catch (Exception)
            {
                preExistingTerm = null;
            }

            // If the matching term is not found, return false... we can't re-use just yet  
            if (preExistingTerm == null)
            {
                return new TryReuseTermResult() { Success = false, UpdatedParser = parser };
            }
            // if the matching term is found re-use, create child terms, and return true  
            else
            {
                // Reuse term
                Term createdTerm = null;
                if (parent is TermSet)
                {
                    createdTerm = ((TermSet)parent).ReuseTerm(preExistingTerm, false);
                }
                else if (parent is Term)
                {
                    createdTerm = ((Term)parent).ReuseTerm(preExistingTerm, false);
                }

                if (modelTerm.IsSourceTerm)
                {
                    preExistingTerm.ReassignSourceTerm(createdTerm);
                }

                // Set labels and shared properties just in case we're on the source term
                if (modelTerm.IsSourceTerm)
                {
                    if (modelTerm.Labels.Any())
                    {
                        CreateTermLabels(modelTerm, termStore, parser, scope, createdTerm);
                    }

                    if (modelTerm.Properties.Any())
                    {
                        SetTermCustomProperties(modelTerm, parser, createdTerm);
                    }
                }

                if (modelTerm.LocalProperties.Any())
                {
                    SetTermLocalCustomProperties(modelTerm, parser, createdTerm);
                }

                termStore.CommitAll();
                web.Context.Load(createdTerm);
                web.Context.ExecuteQueryRetry();

                // Create any child terms
                parser = CreateChildTerms(web, modelTerm, createdTerm, termStore, parser, scope);

                // Return true, because our TryReuseTerm attempt succeeded!
                return new TryReuseTermResult() { Success = true, UpdatedParser = parser };
            }
        }

        private static TokenParser CheckChildTerms(Web web, Model.Term modelTerm, Term parentTerm, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
        {
            if (modelTerm.Terms.Any())
            {
                parentTerm.Context.Load(parentTerm, s => s.Terms.Include(t => t.Id, t => t.Name));
                parentTerm.Context.ExecuteQueryRetry();

                var terms = parentTerm.Terms;

                foreach (var childTerm in modelTerm.Terms)
                {
                    if (terms.Any())
                    {
                        var term = terms.FirstOrDefault(t => t.Id == childTerm.Id);
                        if (term == null)
                        {
                            var normalizedTermName = TaxonomyItem.NormalizeName(web.Context, childTerm.Name);
                            web.Context.ExecuteQueryRetry();

                            term = terms.FirstOrDefault(t => t.Name == normalizedTermName.Value);
                            if (term == null)
                            {
                                var returnTuple = CreateTerm<TermSet>(web, childTerm, parentTerm, termStore, parser, scope);
                                if (returnTuple != null)
                                {
                                    childTerm.Id = returnTuple.Item1;
                                    parser = returnTuple.Item2;
                                }
                            }
                            else
                            {
                                childTerm.Id = term.Id;
                            }
                        }
                        else
                        {
                            childTerm.Id = term.Id;
                        }

                        if (term != null)
                        {
                            parser = CheckChildTerms(web, childTerm, term, termStore, parser, scope);
                        }
                    }
                    else
                    {
                        var returnTuple = CreateTerm<TermSet>(web, childTerm, parentTerm, termStore, parser, scope);
                        if (returnTuple != null)
                        {
                            childTerm.Id = returnTuple.Item1;
                            parser = returnTuple.Item2;
                        }
                    }
                }
            }

            return parser;
        }

        internal class TryReuseTermResult
        {
            public bool Success { get; set; }
            public TokenParser UpdatedParser { get; set; }
        }
    }
}
