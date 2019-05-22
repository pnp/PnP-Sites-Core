using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectTermGroups : ObjectHandlerBase
    {
        private List<TermGroupHelper.ReusedTerm> reusedTerms;

        public override string Name => "Term Groups";

        public override string InternalName => "TermGroups";
        public override TokenParser ProvisionObjects(Web web, Model.ProvisioningTemplate template, TokenParser parser,
            ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                this.reusedTerms = new List<TermGroupHelper.ReusedTerm>();

                TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(web.Context);
                TermStore termStore = null;
                TermGroup siteCollectionTermGroup = null;

                try
                {
                    termStore = taxSession.GetDefaultKeywordsTermStore();
                    web.Context.Load(termStore,
                        ts => ts.Languages,
                        ts => ts.DefaultLanguage,
                        ts => ts.Groups.Include(
                            tg => tg.Name,
                            tg => tg.Id,
                            tg => tg.TermSets.Include(
                                tset => tset.Name,
                                tset => tset.Id)));
                    siteCollectionTermGroup = termStore.GetSiteCollectionGroup((web.Context as ClientContext).Site, false);
                    web.Context.Load(siteCollectionTermGroup);
                    web.Context.ExecuteQueryRetry();
                }
                catch (ServerException)
                {
                    // If the GetDefaultSiteCollectionTermStore method call fails ... raise a specific Warning
                    WriteMessage(CoreResources.Provisioning_ObjectHandlers_TermGroups_Wrong_Configuration, ProvisioningMessageType.Warning);

                    // and exit skipping the current handler
                    return parser;
                }

                SiteCollectionTermGroupNameToken siteCollectionTermGroupNameToken =
                    new SiteCollectionTermGroupNameToken(web);

                foreach (var modelTermGroup in template.TermGroups)
                {
                    this.reusedTerms.AddRange(TermGroupHelper.ProcessGroup(web.Context as ClientContext, taxSession, termStore, modelTermGroup, siteCollectionTermGroup, parser, scope));
                }

                foreach (var reusedTerm in this.reusedTerms)
                {
                    TermGroupHelper.TryReuseTerm(web.Context as ClientContext, reusedTerm.ModelTerm, reusedTerm.Parent, reusedTerm.TermStore, parser, scope);
                }
            }
            return parser;
        }

        //public class ReusedTerm
        //{
        //    public Model.Term ModelTerm { get; set; }
        //    public TaxonomyItem Parent { get; set; }
        //    public TermStore TermStore { get; set; }
        //}

        //private Tuple<Guid, TokenParser> CreateTerm<T>(Web web, Model.Term modelTerm, TaxonomyItem parent,
        //    TermStore termStore, TokenParser parser, PnPMonitoredScope scope) where T : TaxonomyItem
        //{
        //    // If the term is a re-used term and the term is not a source term, skip for now and create later
        //    if (modelTerm.IsReused && !modelTerm.IsSourceTerm)
        //    {
        //        this.reusedTerms.Add(new ReusedTerm()
        //        {
        //            ModelTerm = modelTerm,
        //            Parent = parent,
        //            TermStore = termStore
        //        });
        //        return null;
        //    }

        //    // Create new term
        //    Term term;
        //    if (modelTerm.Id == Guid.Empty)
        //    {
        //        modelTerm.Id = Guid.NewGuid();
        //    }

        //    if (parent is Term)
        //    {
        //        term = ((Term)parent).CreateTerm(parser.ParseString(modelTerm.Name), modelTerm.Language != null && modelTerm.Language != 0 ? modelTerm.Language.Value :  termStore.DefaultLanguage, modelTerm.Id);
        //    }
        //    else
        //    {
        //        term = ((TermSet)parent).CreateTerm(parser.ParseString(modelTerm.Name), modelTerm.Language != null && modelTerm.Language != 0 ? modelTerm.Language.Value : termStore.DefaultLanguage, modelTerm.Id);
        //    }
        //    if (!string.IsNullOrEmpty(modelTerm.Description))
        //    {
        //        term.SetDescription(parser.ParseString(modelTerm.Description), modelTerm.Language != null && modelTerm.Language != 0 ? modelTerm.Language.Value : termStore.DefaultLanguage);
        //    }
        //    if (!string.IsNullOrEmpty(modelTerm.Owner))
        //    {
        //        term.Owner = modelTerm.Owner;
        //    }

        //    term.IsAvailableForTagging = modelTerm.IsAvailableForTagging;

        //    if (modelTerm.Properties.Any() || modelTerm.Labels.Any() || modelTerm.LocalProperties.Any())
        //    {
        //        if (modelTerm.Labels.Any())
        //        {
        //            CreateTermLabels(modelTerm, termStore, parser, scope, term);
        //            //foreach (var label in modelTerm.Labels)
        //            //{
        //            //    if ((label.IsDefaultForLanguage && label.Language != termStore.DefaultLanguage) || label.IsDefaultForLanguage == false)
        //            //    {
        //            //        term.CreateLabel(parser.ParseString(label.Value), label.Language, label.IsDefaultForLanguage);
        //            //    }
        //            //    else
        //            //    {
        //            //        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language);
        //            //        WriteWarning(string.Format(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language), ProvisioningMessageType.Warning);
        //            //    }
        //            //}
        //        }

        //        if (modelTerm.Properties.Any())
        //        {
        //            SetTermCustomProperties(modelTerm, parser, term);
        //            //foreach (var property in modelTerm.Properties)
        //            //{
        //            //    term.SetCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
        //            //}
        //        }
        //        if (modelTerm.LocalProperties.Any())
        //        {
        //            SetTermLocalCustomProperties(modelTerm, parser, term);
        //            //foreach (var property in modelTerm.LocalProperties)
        //            //{
        //            //    term.SetLocalCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
        //            //}
        //        }
        //    }

        //    termStore.CommitAll();

        //    web.Context.Load(term);
        //    web.Context.ExecuteQueryRetry();

        //    // Deprecate term if needed
        //    if (modelTerm.IsDeprecated != term.IsDeprecated)
        //    {
        //        term.Deprecate(modelTerm.IsDeprecated);
        //        web.Context.ExecuteQueryRetry();
        //    }


        //    parser = this.CreateChildTerms(web, modelTerm, term, termStore, parser, scope);
        //    return Tuple.Create(modelTerm.Id, parser);
        //}


        //private void CreateTermLabels(Model.Term modelTerm, TermStore termStore, TokenParser parser, PnPMonitoredScope scope, Term term)
        //{
        //    foreach (var label in modelTerm.Labels)
        //    {
        //        if (((label.IsDefaultForLanguage && label.Language != termStore.DefaultLanguage) || label.IsDefaultForLanguage == false) && termStore.Languages.Contains(label.Language))
        //        {
        //            term.CreateLabel(parser.ParseString(label.Value), label.Language, label.IsDefaultForLanguage);
        //        }
        //        else
        //        {
        //            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language);
        //            WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language), ProvisioningMessageType.Warning);
        //        }
        //    }
        //}

        //private static void SetTermCustomProperties(Model.Term modelTerm, TokenParser parser, Term term)
        //{
        //    foreach (var property in modelTerm.Properties)
        //    {
        //        term.SetCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
        //    }
        //}

        //private static void SetTermLocalCustomProperties(Model.Term modelTerm, TokenParser parser, Term term)
        //{
        //    foreach (var property in modelTerm.LocalProperties)
        //    {
        //        term.SetLocalCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
        //    }
        //}

        // /// <summary>
        // /// Creates child terms for the current model term if any exist
        // /// </summary>
        // /// <param name="web"></param>
        // /// <param name="modelTerm"></param>
        // /// <param name="term"></param>
        // /// <param name="termStore"></param>
        // /// <param name="parser"></param>
        // /// <param name="scope"></param>
        // /// <returns>Updated parser object</returns>
        //private TokenParser CreateChildTerms(Web web, Model.Term modelTerm, Term term, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
        //{
        //    if (modelTerm.Terms.Any())
        //    {
        //        foreach (var modelTermTerm in modelTerm.Terms)
        //        {
        //            web.Context.Load(term.Terms);
        //            web.Context.ExecuteQueryRetry();
        //            var termTerms = term.Terms;
        //            if (termTerms.Any())
        //            {
        //                var termTerm = termTerms.FirstOrDefault(t => t.Id == modelTermTerm.Id);
        //                if (termTerm == null)
        //                {
        //                    termTerm = termTerms.FirstOrDefault(t => t.Name == modelTermTerm.Name);
        //                    if (termTerm == null)
        //                    {
        //                        var returnTuple = CreateTerm<Term>(web, modelTermTerm, term, termStore, parser, scope);
        //                        if (returnTuple != null)
        //                        {
        //                            modelTermTerm.Id = returnTuple.Item1;
        //                            parser = returnTuple.Item2;
        //                        }
        //                    }
        //                    else
        //                    {
        //                        modelTermTerm.Id = termTerm.Id;
        //                    }
        //                }
        //                else
        //                {
        //                    modelTermTerm.Id = termTerm.Id;
        //                }
        //            }
        //            else
        //            {
        //                var returnTuple = CreateTerm<Term>(web, modelTermTerm, term, termStore, parser, scope);
        //                if (returnTuple != null)
        //                {
        //                    modelTermTerm.Id = returnTuple.Item1;
        //                    parser = returnTuple.Item2;
        //                }
        //            }
        //        }
        //        if (modelTerm.Terms.Any(t => t.CustomSortOrder > -1))
        //        {
        //            var sortedTerms = modelTerm.Terms.OrderBy(t => t.CustomSortOrder);

        //            var customSortString = sortedTerms.Aggregate(string.Empty, (a, i) => a + i.Id.ToString() + ":");
        //            customSortString = customSortString.TrimEnd(new[] { ':' });

        //            term.CustomSortOrder = customSortString;
        //            termStore.CommitAll();
        //        }
        //    }

        //    return parser;
        //}

        // /// <summary>
        // /// Attempts to reuse the model term. If the term does not yet exists it will return
        // /// false for the first part of the the return tuple. this will notify the system
        // /// that the term should be created instead of re-used.
        // /// </summary>
        // /// <param name="web"></param>
        // /// <param name="modelTerm"></param>
        // /// <param name="parent"></param>
        // /// <param name="termStore"></param>
        // /// <param name="parser"></param>
        // /// <param name="scope"></param>
        // /// <returns></returns>


        //private TokenParser CheckChildTerms(Web web, Model.Term modelTerm, Term parentTerm, TermStore termStore, TokenParser parser, PnPMonitoredScope scope)
        //{
        //    if (modelTerm.Terms.Any())
        //    {
        //        parentTerm.Context.Load(parentTerm, s => s.Terms.Include(t => t.Id, t => t.Name));
        //        parentTerm.Context.ExecuteQueryRetry();

        //        var terms = parentTerm.Terms;

        //        foreach (var childTerm in modelTerm.Terms)
        //        {
        //            if (terms.Any())
        //            {
        //                var term = terms.FirstOrDefault(t => t.Id == childTerm.Id);
        //                if (term == null)
        //                {
        //                    var normalizedTermName = TaxonomyItem.NormalizeName(web.Context, childTerm.Name);
        //                    web.Context.ExecuteQueryRetry();

        //                    term = terms.FirstOrDefault(t => t.Name == normalizedTermName.Value);
        //                    if (term == null)
        //                    {
        //                        var returnTuple = TermGroupHelper.CreateTerm<TermSet>(web, childTerm, parentTerm, termStore, parser, scope);
        //                        if (returnTuple != null)
        //                        {
        //                            childTerm.Id = returnTuple.Item1;
        //                            parser = returnTuple.Item2;
        //                            this.reusedTerms.AddRange(returnTuple.Item3);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        childTerm.Id = term.Id;
        //                    }
        //                }
        //                else
        //                {
        //                    childTerm.Id = term.Id;
        //                }

        //                if (term != null)
        //                {
        //                    parser = CheckChildTerms(web, childTerm, term, termStore, parser, scope);
        //                }
        //            }
        //            else
        //            {
        //                var returnTuple = TermGroupHelper.CreateTerm<TermSet>(web, childTerm, parentTerm, termStore, parser, scope);
        //                if (returnTuple != null)
        //                {
        //                    childTerm.Id = returnTuple.Item1;
        //                    parser = returnTuple.Item2;
        //                    this.reusedTerms.AddRange(returnTuple.Item3);
        //                }
        //            }
        //        }
        //    }

        //    return parser;
        //}

        private class TryReuseTermResult
        {
            public bool Success { get; set; }
            public TokenParser UpdatedParser { get; set; }
        }

        public override Model.ProvisioningTemplate ExtractObjects(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (creationInfo.IncludeSiteCollectionTermGroup || creationInfo.IncludeAllTermGroups)
                {
                    // Find the site collection termgroup, if any
                    TaxonomySession session = TaxonomySession.GetTaxonomySession(web.Context);
                    TermStore termStore = null;

                    try
                    {
                        termStore = session.GetDefaultSiteCollectionTermStore();
                        web.Context.Load(termStore, t => t.Id, t => t.DefaultLanguage, t => t.OrphanedTermsTermSet);
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (ServerException)
                    {
                        // Skip the exception and go to the next check
                    }

                    if (null == termStore || termStore.ServerObjectIsNull())
                    {
                        // If the GetDefaultSiteCollectionTermStore method call fails ... raise a specific Warning
                        WriteMessage(CoreResources.Provisioning_ObjectHandlers_TermGroups_Wrong_Configuration, ProvisioningMessageType.Warning);

                        // and exit skipping the current handler
                        return template;
                    }

                    var orphanedTermsTermSetId = default(Guid);
                    if (!termStore.OrphanedTermsTermSet.ServerObjectIsNull())
                    {
                        termStore.OrphanedTermsTermSet.EnsureProperty(ts => ts.Id);
                        orphanedTermsTermSetId = termStore.OrphanedTermsTermSet.Id;
                        if (termStore.ServerObjectIsNull.Value)
                        {
                            termStore = session.GetDefaultKeywordsTermStore();
                            web.Context.Load(termStore, t => t.Id, t => t.DefaultLanguage);
                            web.Context.ExecuteQueryRetry();
                        }
                    }

                    var propertyBagKey = $"SiteCollectionGroupId{termStore.Id}";

                    // Ensure to grab the property from the rootweb
                    var site = (web.Context as ClientContext).Site;
                    web.Context.Load(site, s => s.RootWeb);
                    web.Context.ExecuteQueryRetry();

                    var siteCollectionTermGroupId = site.RootWeb.GetPropertyBagValueString(propertyBagKey, "");

                    Guid termGroupGuid;
                    Guid.TryParse(siteCollectionTermGroupId, out termGroupGuid);

                    List<TermGroup> termGroups = new List<TermGroup>();
                    if (creationInfo.IncludeAllTermGroups)
                    {
                        web.Context.Load(termStore.Groups, groups => groups.Include(tg => tg.Name,
                            tg => tg.Id,
                            tg => tg.Description,
                            tg => tg.TermSets.IncludeWithDefaultProperties(ts => ts.CustomSortOrder)));
                        web.Context.ExecuteQueryRetry();
                        termGroups = termStore.Groups.ToList();
                    }
                    else
                    {
                        if (termGroupGuid != Guid.Empty)
                        {
                            var termGroup = termStore.GetGroup(termGroupGuid);
                            web.Context.Load(termGroup,
                                tg => tg.Name,
                                tg => tg.Id,
                                tg => tg.Description,
                                tg => tg.TermSets.IncludeWithDefaultProperties(ts => ts.Description, ts => ts.CustomSortOrder));

                            web.Context.ExecuteQueryRetry();

                            termGroups = new List<TermGroup>() { termGroup };
                        }
                    }

                    foreach (var termGroup in termGroups)
                    {
                        Boolean isSiteCollectionTermGroup = termGroupGuid != Guid.Empty && termGroup.Id == termGroupGuid;

                        var modelTermGroup = new Model.TermGroup
                        {
                            Name = isSiteCollectionTermGroup ? "{sitecollectiontermgroupname}" : termGroup.Name,
                            Id = isSiteCollectionTermGroup ? Guid.Empty : termGroup.Id,
                            Description = termGroup.Description,
                            SiteCollectionTermGroup = isSiteCollectionTermGroup
                        };

#if !ONPREMISES

                        // If we need to include TermGroups security
                        if (creationInfo.IncludeTermGroupsSecurity)
                        {
                            termGroup.EnsureProperties(tg => tg.ContributorPrincipalNames, tg => tg.GroupManagerPrincipalNames);

                            // Extract the TermGroup contributors
                            modelTermGroup.Contributors.AddRange(
                                from c in termGroup.ContributorPrincipalNames
                                select new Model.User { Name = c });

                            // Extract the TermGroup managers
                            modelTermGroup.Managers.AddRange(
                                from m in termGroup.GroupManagerPrincipalNames
                                select new Model.User { Name = m });
                        }

#endif

                        web.EnsureProperty(w => w.Url);

                        foreach (var termSet in termGroup.TermSets)
                        {
                            // Do not include the orphan term set
                            if (termSet.Id == orphanedTermsTermSetId) continue;

                            // Extract all other term sets
                            var modelTermSet = new Model.TermSet();
                            modelTermSet.Name = termSet.Name;
                            if (!isSiteCollectionTermGroup)
                            {
                                modelTermSet.Id = termSet.Id;
                            }
                            modelTermSet.IsAvailableForTagging = termSet.IsAvailableForTagging;
                            modelTermSet.IsOpenForTermCreation = termSet.IsOpenForTermCreation;
                            modelTermSet.Description = termSet.Description;
                            modelTermSet.Terms.AddRange(GetTerms<TermSet>(web.Context, termSet, termStore.DefaultLanguage, isSiteCollectionTermGroup));
                            foreach (var property in termSet.CustomProperties)
                            {
                                if (property.Key.Equals("_Sys_Nav_AttachedWeb_SiteId", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    modelTermSet.Properties.Add(property.Key, "{sitecollectionid}");
                                }
                                else if (property.Key.Equals("_Sys_Nav_AttachedWeb_WebId", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    modelTermSet.Properties.Add(property.Key, "{siteid}");
                                }
                                else
                                {
                                    modelTermSet.Properties.Add(property.Key, Tokenize(property.Value, web.Url, web));
                                }
                            }
                            modelTermGroup.TermSets.Add(modelTermSet);
                        }

                        template.TermGroups.Add(modelTermGroup);
                    }
                }
            }
            return template;
        }

        private List<Model.Term> GetTerms<T>(ClientRuntimeContext context, TaxonomyItem parent, int defaultLanguage, Boolean isSiteCollectionTermGroup = false)
        {
            List<Model.Term> termsToReturn = new List<Model.Term>();
            TermCollection terms;
            var customSortOrder = string.Empty;
            if (parent is TermSet)
            {
                terms = ((TermSet)parent).Terms;
                customSortOrder = ((TermSet)parent).CustomSortOrder;
            }
            else
            {
                terms = ((Term)parent).Terms;
                customSortOrder = ((Term)parent).CustomSortOrder;
            }
            context.Load(terms, tms => tms.IncludeWithDefaultProperties(t => t.Labels, t => t.CustomSortOrder,
                t => t.IsReused, t => t.IsSourceTerm, t => t.SourceTerm, t => t.IsDeprecated, t => t.Description, t => t.Owner));
            context.ExecuteQueryRetry();

            foreach (var term in terms)
            {
                var modelTerm = new Model.Term();
                if (!isSiteCollectionTermGroup || term.IsReused)
                {
                    modelTerm.Id = term.Id;
                }
                modelTerm.Name = term.Name;
                modelTerm.IsAvailableForTagging = term.IsAvailableForTagging;
                modelTerm.IsReused = term.IsReused;
                modelTerm.IsSourceTerm = term.IsSourceTerm;
                modelTerm.SourceTermId = (term.SourceTerm != null) ? term.SourceTerm.Id : Guid.Empty;
                modelTerm.IsDeprecated = term.IsDeprecated;
                modelTerm.Description = term.Description;
                modelTerm.Owner = term.Owner;

                if ((!term.IsReused || term.IsSourceTerm) && term.Labels.Any())
                {
                    foreach (var label in term.Labels)
                    {
                        if ((label.Language == defaultLanguage && label.Value != term.Name) || label.Language != defaultLanguage)
                        {
                            var modelLabel = new Model.TermLabel();
                            modelLabel.IsDefaultForLanguage = label.IsDefaultForLanguage;
                            modelLabel.Value = label.Value;
                            modelLabel.Language = label.Language;

                            modelTerm.Labels.Add(modelLabel);
                        }
                    }
                }
                //else
                //{
                //    foreach (var label in term.Labels)
                //    {
                //        var modelLabel = new Model.TermLabel();
                //        modelLabel.IsDefaultForLanguage = label.IsDefaultForLanguage;
                //        modelLabel.Value = label.Value;
                //        modelLabel.Language = label.Language;

                //        modelTerm.Labels.Add(modelLabel);
                //    }
                //}

                foreach (var localProperty in term.LocalCustomProperties)
                {
                    modelTerm.LocalProperties.Add(localProperty.Key, localProperty.Value);
                }

                // Shared Properties have to be extracted just for source terms or not reused terms
                if (!term.IsReused || term.IsSourceTerm)
                {
                    foreach (var customProperty in term.CustomProperties)
                    {
                        modelTerm.Properties.Add(customProperty.Key, customProperty.Value);
                    }
                }

                if (term.TermsCount > 0)
                {
                    modelTerm.Terms.AddRange(GetTerms<Term>(context, term, defaultLanguage, isSiteCollectionTermGroup));
                }
                termsToReturn.Add(modelTerm);

                if (!string.IsNullOrEmpty(customSortOrder))
                {
                    var sortOrder = customSortOrder.Split(new[] { ':' }).ToList();

                    var currentTermIndex = sortOrder.Where(i => new Guid(i) == term.Id).FirstOrDefault();
                    modelTerm.CustomSortOrder = sortOrder.IndexOf(currentTermIndex) + 1;

                }
            }
            termsToReturn = termsToReturn.OrderBy(t => t.CustomSortOrder).ToList();

            return termsToReturn;
        }


        public override bool WillProvision(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.TermGroups.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = creationInfo.IncludeSiteCollectionTermGroup || creationInfo.IncludeAllTermGroups;
            }
            return _willExtract.Value;
        }
    }
}
