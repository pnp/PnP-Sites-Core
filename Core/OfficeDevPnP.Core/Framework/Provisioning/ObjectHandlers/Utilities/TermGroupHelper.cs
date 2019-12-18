using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    internal static class TermGroupHelper
    {
        internal static List<ReusedTerm> ProcessGroup(ClientContext context, TaxonomySession session, TermStore termStore, Model.TermGroup modelTermGroup, TermGroup siteCollectionTermGroup, TokenParser parser, PnPMonitoredScope scope)
        {
            List<ReusedTerm> reusedTerms = new List<ReusedTerm>();

            SiteCollectionTermGroupNameToken siteCollectionTermGroupNameToken =
              new SiteCollectionTermGroupNameToken(context.Web);

            #region Group

            var newGroup = false;

            var modelGroupName = parser.ParseString(modelTermGroup.Name);

            var normalizedGroupName = TaxonomyItem.NormalizeName(context, modelGroupName);
            context.ExecuteQueryRetry();

            TermGroup group = termStore.Groups.FirstOrDefault(
                g => g.Id == modelTermGroup.Id || g.Name == normalizedGroupName.Value);
            if (group == null)
            {
                var parsedDescription = parser.ParseString(modelTermGroup.Description);

                if (modelTermGroup.Name == "Site Collection" ||
                    modelGroupName == siteCollectionTermGroupNameToken.GetReplaceValue() ||
                    modelTermGroup.SiteCollectionTermGroup)
                {
                    var site = (context as ClientContext).Site;
                    group = termStore.GetSiteCollectionGroup(site, true);
                    context.Load(group, g => g.Name, g => g.Id, g => g.TermSets.Include(
                        tset => tset.Name,
                        tset => tset.Id));
                    context.ExecuteQueryRetry();
                }
                else
                {
                    group = termStore.Groups.FirstOrDefault(g => g.Name == normalizedGroupName.Value);

                    if (group == null)
                    {
                        if (modelTermGroup.Id == Guid.Empty)
                        {
                            modelTermGroup.Id = Guid.NewGuid();
                        }
                        group = termStore.CreateGroup(modelGroupName, modelTermGroup.Id);

                        group.Description = parsedDescription;

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

                        termStore.CommitAll();
                        context.Load(group);
                        context.Load(termStore);
                        context.ExecuteQueryRetry();

                        newGroup = true;

                    }
                }

            }
            #endregion

            session.UpdateCache();
            session.Context.ExecuteQueryRetry();
            #region TermSets

            foreach (var modelTermSet in modelTermGroup.TermSets)
            {
                TermSet set = null;
                var newTermSet = false;

                var normalizedTermSetName = TaxonomyItem.NormalizeName(context, parser.ParseString(modelTermSet.Name));
                context.ExecuteQueryRetry();
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
                    else
                    {
                        if (CheckIfTermSetIdIsUnique(termStore, modelTermSet.Id) == false)
                        {
                            throw new Exception($"Termset ID {modelTermSet.Id} is already present in termstore");
                        }
                    }
                    var termSetLanguage = modelTermSet.Language.HasValue ? modelTermSet.Language.Value : termStore.DefaultLanguage;
                    set = group.CreateTermSet(normalizedTermSetName.Value, modelTermSet.Id, termSetLanguage);
                    parser.AddToken(new TermSetIdToken(context.Web, group.Name, normalizedTermSetName.Value, modelTermSet.Id));
                    if (siteCollectionTermGroup != null && !siteCollectionTermGroup.ServerObjectIsNull.Value)
                    {
                        if (group.Name == siteCollectionTermGroup.Name)
                        {
                            parser.AddToken((new SiteCollectionTermSetIdToken(context.Web, normalizedTermSetName.Value, modelTermSet.Id)));
                        }
                    }
                    newTermSet = true;
                    if (!string.IsNullOrEmpty(modelTermSet.Description))
                    {
                        set.Description = parser.ParseString(modelTermSet.Description);
                    }
                    set.IsOpenForTermCreation = modelTermSet.IsOpenForTermCreation;
                    set.IsAvailableForTagging = modelTermSet.IsAvailableForTagging;
                    foreach (var property in modelTermSet.Properties)
                    {
                        set.SetCustomProperty(property.Key, parser.ParseString(property.Value));
                    }
                    if (modelTermSet.Owner != null)
                    {
                        set.Owner = parser.ParseString(modelTermSet.Owner);
                    }
                    termStore.CommitAll();
                    context.Load(set);
                    context.ExecuteQueryRetry();
                }

                context.Load(set, s => s.Terms.Include(t => t.Id, t => t.Name));
                context.ExecuteQueryRetry();
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
                                var normalizedTermName = TaxonomyItem.NormalizeName(context, parser.ParseString(modelTerm.Name));
                                context.ExecuteQueryRetry();

                                term = terms.FirstOrDefault(t => t.Name == normalizedTermName.Value);
                                if (term == null)
                                {
                                    var returnTuple = CreateTerm(context, modelTerm, set, termStore, parser, scope);
                                    if (returnTuple != null)
                                    {
                                        modelTerm.Id = returnTuple.Item1;
                                        parser = returnTuple.Item2;
                                    }
                                    reusedTerms.AddRange(returnTuple.Item3);
                                }
                                else
                                {
                                    // todo: add handling for reused term?
                                    modelTerm.Id = term.Id;
                                }
                            }
                            else
                            {
                                // todo: add handling for reused term?
                                modelTerm.Id = term.Id;
                            }

                            if (term != null)
                            {
                                CheckChildTerms(context, modelTerm, term, termStore, parser, scope);
                            }
                        }
                        else
                        {
                            var returnTuple = CreateTerm(context, modelTerm, set, termStore, parser, scope);
                            if (returnTuple != null)
                            {
                                modelTerm.Id = returnTuple.Item1;
                                parser = returnTuple.Item2;
                            }
                            reusedTerms.AddRange(returnTuple.Item3);
                        }
                    }
                    else
                    {
                        var returnTuple = CreateTerm(context, modelTerm, set, termStore, parser, scope);
                        if (returnTuple != null)
                        {
                            modelTerm.Id = returnTuple.Item1;
                            parser = returnTuple.Item2;
                        }
                        reusedTerms.AddRange(returnTuple.Item3);
                    }
                }

                // do we need custom sorting?
                if (modelTermSet.Terms.Any(t => t.CustomSortOrder > 0))
                {
                    var sortedTerms = modelTermSet.Terms.OrderBy(t => t.CustomSortOrder);

                    var customSortString = sortedTerms.Aggregate(string.Empty,
                        (a, i) => a + i.Id.ToString() + ":");
                    customSortString = customSortString.TrimEnd(new[] { ':' });

                    set.CustomSortOrder = customSortString;
                    termStore.CommitAll();
                    context.ExecuteQueryRetry();
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

        internal static bool CheckIfTermSetIdIsUnique(TermStore store, Guid id)
        {
            var existingTermSet = store.GetTermSet(id);
            store.Context.Load(existingTermSet);
            store.Context.ExecuteQueryRetry();

            return existingTermSet.ServerObjectIsNull == true;
        }

        internal static bool CheckIfTermIdIsUnique(TermStore store, Guid id)
        {
            var existingTerm = store.GetTerm(id);
            store.Context.Load(existingTerm);
            store.Context.ExecuteQueryRetry();

            return existingTerm.ServerObjectIsNull == true;
        }

        internal static Tuple<Guid, TokenParser, List<ReusedTerm>> CreateTerm(ClientContext context, Model.Term modelTerm, TaxonomyItem parent,
           TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
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
                return Tuple.Create(modelTerm.Id, parser, reusedTerms);
            }

            // Create new term
            Term term;
            if (modelTerm.Id == Guid.Empty)
            {
                modelTerm.Id = Guid.NewGuid();
            }
            else
            {
                if (CheckIfTermIdIsUnique(termStore, modelTerm.Id) == false)
                {
                    throw new Exception($"Term ID {modelTerm.Id} is already present in termstore");
                }
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

            context.Load(term);
            context.ExecuteQueryRetry();

            // Deprecate term if needed
            if (modelTerm.IsDeprecated != term.IsDeprecated)
            {
                term.Deprecate(modelTerm.IsDeprecated);
                context.ExecuteQueryRetry();
            }


            parser = CreateChildTerms(context, modelTerm, term, termStore, parser, scope);
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
        /// <param name="context"></param>
        /// <param name="modelTerm"></param>
        /// <param name="term"></param>
        /// <param name="termStore"></param>
        /// <param name="parser"></param>
        /// <param name="scope"></param>
        /// <returns>Updated parser object</returns>
        private static TokenParser CreateChildTerms(ClientContext context, Model.Term modelTerm, Term term, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
        {
            if (modelTerm.Terms.Any())
            {
                foreach (var modelTermTerm in modelTerm.Terms)
                {
                    context.Load(term.Terms);
                    context.ExecuteQueryRetry();
                    var termTerms = term.Terms;
                    if (termTerms.Any())
                    {
                        var termTerm = termTerms.FirstOrDefault(t => t.Id == modelTermTerm.Id);
                        if (termTerm == null)
                        {
                            termTerm = termTerms.FirstOrDefault(t => t.Name == modelTermTerm.Name);
                            if (termTerm == null)
                            {
                                var returnTuple = CreateTerm(context, modelTermTerm, term, termStore, parser, scope);
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
                        var returnTuple = CreateTerm(context, modelTermTerm, term, termStore, parser, scope);
                        if (returnTuple != null)
                        {
                            modelTermTerm.Id = returnTuple.Item1;
                            parser = returnTuple.Item2;
                        }
                    }
                }
                if (modelTerm.Terms.Any(t => t.CustomSortOrder > 0))
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
        /// <param name="context"></param>
        /// <param name="modelTerm"></param>
        /// <param name="parent"></param>
        /// <param name="termStore"></param>
        /// <param name="parser"></param>
        /// <param name="scope"></param>
        /// <returns></returns>
        internal static TryReuseTermResult TryReuseTerm(ClientContext context, Model.Term modelTerm, TaxonomyItem parent, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
        {
            if (!modelTerm.IsReused) return new TryReuseTermResult() { Success = false, UpdatedParser = parser };
            if (modelTerm.Id == Guid.Empty) return new TryReuseTermResult() { Success = false, UpdatedParser = parser };

            // Since we're reusing terms ensure the previous terms are committed
            termStore.CommitAll();
            context.ExecuteQueryRetry();

            // Try to retrieve a matching term from the website also marked from re-use.  
            var taxonomySession = TaxonomySession.GetTaxonomySession(context);
            context.Load(taxonomySession);
            context.ExecuteQueryRetry();

            if (taxonomySession.ServerObjectIsNull())
            {
                return new TryReuseTermResult() { Success = false, UpdatedParser = parser };
            }

            var freshTermStore = taxonomySession.GetDefaultKeywordsTermStore();
            Term preExistingTerm = freshTermStore.GetTerm(modelTerm.Id);

            try
            {
                context.Load(preExistingTerm);
                context.ExecuteQueryRetry();

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
                context.Load(createdTerm);
                context.ExecuteQueryRetry();

                // Create any child terms
                parser = CreateChildTerms(context, modelTerm, createdTerm, termStore, parser, scope);

                // Return true, because our TryReuseTerm attempt succeeded!
                return new TryReuseTermResult() { Success = true, UpdatedParser = parser };
            }
        }

        private static TokenParser CheckChildTerms(ClientContext context, Model.Term modelTerm, Term parentTerm, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
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
                            var normalizedTermName = TaxonomyItem.NormalizeName(context, childTerm.Name);
                            context.ExecuteQueryRetry();

                            term = terms.FirstOrDefault(t => t.Name == normalizedTermName.Value);
                            if (term == null)
                            {
                                var returnTuple = CreateTerm(context, childTerm, parentTerm, termStore, parser, scope);
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
                            parser = CheckChildTerms(context, childTerm, term, termStore, parser, scope);
                        }
                    }
                    else
                    {
                        var returnTuple = CreateTerm(context, childTerm, parentTerm, termStore, parser, scope);
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
