using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectTermGroups : ObjectHandlerBase
    {

        public override string Name
        {
            get { return "Term Groups"; }
        }
        public override TokenParser ProvisionObjects(Web web, Model.ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {

                var termStore = GetTermStoreFor(web);

                var termGroupsContext = new TermGroupsContext
                {
                    TermStore = termStore,
                    Parser = parser,
                    Scope = scope
                };

                foreach (var termGroup in template.TermGroups)
                {
                    var provisionedTermGroup = EnsureTermGroup(termGroup, termGroupsContext);
                    foreach (var termSet in termGroup.TermSets)
                    {
                        var provisionedTermSet = EnsureTermSet(termSet, provisionedTermGroup, termGroupsContext);
                        EnsureTerms(termSet.Terms, provisionedTermSet, termGroupsContext);
                    }
                }
            }
            return parser;
        }

        private void EnsureTerms(Model.TermCollection terms, TermSetItem parent, TermGroupsContext termGroupsContext)
        {
            var termStore = termGroupsContext.TermStore;
            var parser = termGroupsContext.Parser;
            foreach (var term in terms)
            {
                if (term.Id == Guid.Empty)
                {
                    term.Id = Guid.NewGuid();
                }
                var provisionedItem = termStore.GetTerm(term.Id);
                termStore.Context.Load(provisionedItem);
                termStore.Context.ExecuteQuery();
                

                if (provisionedItem.ServerObjectIsNull == true)
                {
                    provisionedItem = parent.CreateTerm(parser.ParseString(term.Name), term.Language ?? termStore.DefaultLanguage, term.Id);
                    provisionedItem.SetDescription(term.Description??string.Empty, term.Language ?? termStore.DefaultLanguage);
                    if (!string.IsNullOrEmpty(term.Owner))
                    {
                        provisionedItem.Owner = term.Owner;
                    }
                    provisionedItem.IsAvailableForTagging = term.IsAvailableForTagging;

                    foreach (var label in term.Labels)
                    {
                        if ((label.IsDefaultForLanguage && label.Language != termStore.DefaultLanguage) || label.IsDefaultForLanguage == false)
                        {
                            var l = provisionedItem.CreateLabel(parser.ParseString(label.Value), label.Language, label.IsDefaultForLanguage);
                        }
                        else
                        {
                            termGroupsContext.Scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language);
                            WriteWarning(string.Format(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language), ProvisioningMessageType.Warning);
                        }
                    }

                    foreach (var property in term.Properties)
                    {
                        provisionedItem.SetCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
                    }

                    foreach (var property in term.LocalProperties)
                    {
                        provisionedItem.SetLocalCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
                    }
                    termStore.CommitAll();
                    termStore.Context.ExecuteQuery();
                }
                EnsureTerms(term.Terms,provisionedItem,termGroupsContext);

            }
        }

        private TermSet EnsureTermSet(Model.TermSet termSet, TermGroup provisionedTermGroup, TermGroupsContext termGroupsContext)
        {
            var termStore = termGroupsContext.TermStore;
            TermSet provisionedTermSet = null;
            try
            {
                provisionedTermSet = provisionedTermGroup.TermSets.FirstOrDefault(ts => ts.Id == termSet.Id || ts.Name == termSet.Name);
            }
            catch (CollectionNotInitializedException ex)
            {
                // no termsets will yield a CollectionNotInitializedException
            }

            if (provisionedTermSet == null)
            {
                if (termSet.Id == Guid.Empty)
                {
                    termSet.Id = Guid.NewGuid();
                }
                var parser = termGroupsContext.Parser;
                provisionedTermSet = provisionedTermGroup.CreateTermSet(parser.ParseString(termSet.Name), termSet.Id, termSet.Language ?? termStore.DefaultLanguage);
                parser.AddToken(new TermSetIdToken(termStore.Context.GetSiteCollectionContext().Web, termSet.Name, termSet.Name, termSet.Id));
                provisionedTermSet.IsOpenForTermCreation = termSet.IsOpenForTermCreation;
                provisionedTermSet.IsAvailableForTagging = termSet.IsAvailableForTagging;
                foreach (var property in termSet.Properties)
                {
                    provisionedTermSet.SetCustomProperty(property.Key, property.Value);
                }
                if (termSet.Owner != null)
                {
                    provisionedTermSet.Owner = termSet.Owner;
                }
                termStore.CommitAll();
                termStore.Context.Load(provisionedTermSet);
                termStore.Context.ExecuteQueryRetry();

            }
            termStore.Context.Load(provisionedTermSet, s => s.Terms.Include(t => t.Id, t => t.Name));
            termStore.Context.ExecuteQueryRetry();
            return provisionedTermSet;
        }

        private TermGroup EnsureTermGroup(Model.TermGroup termGroup, TermGroupsContext termGroupsCtx)
        {
            var termStore = termGroupsCtx.TermStore;

            TermGroup group = termStore.Groups.FirstOrDefault(g => g.Id == termGroup.Id || g.Name == termGroup.Name);
            if (group == null)
            {
                if (termGroup.Name == "Site Collection")
                {
                    var site = (termStore.Context as ClientContext).Site;
                    group = termStore.GetSiteCollectionGroup(site, true);
                    site.Context.Load(group, g => g.Name, g => g.Id, g => g.TermSets.Include(
                        tset => tset.Name,
                        tset => tset.Id));
                    site.Context.ExecuteQueryRetry();
                }
                else
                {
                    var parsedGroupName = termGroupsCtx.Parser.ParseString(termGroup.Name);
                    group = termStore.Groups.FirstOrDefault(g => g.Name == parsedGroupName);

                    if (group == null)
                    {
                        if (termGroup.Id == Guid.Empty)
                        {
                            termGroup.Id = Guid.NewGuid();
                        }
                        group = termStore.CreateGroup(parsedGroupName, termGroup.Id);

                        group.Description = termGroup.Description;

                        termStore.CommitAll();
                        termStore.Context.Load(group);
                        termStore.Context.ExecuteQueryRetry();

                    }
                }
            }
            return group;
        }


        private static TermStore GetTermStoreFor(Web web)
        {
            TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(web.Context);

            var termStore = taxSession.GetDefaultKeywordsTermStore();

            web.Context.Load(termStore,
                ts => ts.DefaultLanguage,
                ts => ts.Groups.Include(
                    tg => tg.Name,
                    tg => tg.Id,
                    tg => tg.TermSets.Include(
                        tset => tset.Name,
                        tset => tset.Id)));
            web.Context.ExecuteQueryRetry();
            return termStore;
        }

       

        public override Model.ProvisioningTemplate ExtractObjects(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (creationInfo.IncludeSiteCollectionTermGroup || creationInfo.IncludeAllTermGroups)
                {
                    // Find the site collection termgroup, if any
                    TaxonomySession session = TaxonomySession.GetTaxonomySession(web.Context);
                    var termStore = session.GetDefaultSiteCollectionTermStore();
					web.Context.Load(termStore, t => t.Id, t => t.DefaultLanguage, t => t.OrphanedTermsTermSet);
					web.Context.ExecuteQueryRetry();

                    var orphanedTermsTermSetId = termStore.OrphanedTermsTermSet.Id;
					if (termStore.ServerObjectIsNull.Value)
					{
						termStore = session.GetDefaultKeywordsTermStore();
						web.Context.Load(termStore, t => t.Id, t => t.DefaultLanguage);
						web.Context.ExecuteQueryRetry();
					}

                    var propertyBagKey = string.Format("SiteCollectionGroupId{0}", termStore.Id);

                    var siteCollectionTermGroupId = web.GetPropertyBagValueString(propertyBagKey, "");

                    Guid termGroupGuid = Guid.Empty;
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
                                tg => tg.TermSets.IncludeWithDefaultProperties(ts => ts.CustomSortOrder));

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
                            Description = termGroup.Description
                        };

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
                                modelTermSet.Properties.Add(property.Key, property.Value);
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
            TermCollection terms = null;
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
            context.Load(terms, tms => tms.IncludeWithDefaultProperties(t => t.Labels, t => t.CustomSortOrder));
            context.ExecuteQueryRetry();

            foreach (var term in terms)
            {
                var modelTerm = new Model.Term();
                if (!isSiteCollectionTermGroup)
                {
                    modelTerm.Id = term.Id;
                }
                modelTerm.Name = term.Name;
                modelTerm.IsAvailableForTagging = term.IsAvailableForTagging;


                if (term.Labels.Any())
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

                foreach (var localProperty in term.LocalCustomProperties)
                {
                    modelTerm.LocalProperties.Add(localProperty.Key, localProperty.Value);
                }

                foreach (var customProperty in term.CustomProperties)
                {
                    modelTerm.Properties.Add(customProperty.Key, customProperty.Value);
                }
                if (term.TermsCount > 0)
                {
                    modelTerm.Terms.AddRange(GetTerms<Term>(context, term, defaultLanguage, isSiteCollectionTermGroup));
                }
                termsToReturn.Add(modelTerm);
            }
            if (!string.IsNullOrEmpty(customSortOrder))
            {
                int count = 1;
                foreach (var id in customSortOrder.Split(new[] { ':' }))
                {
                    var term = termsToReturn.FirstOrDefault(t => t.Id == Guid.Parse(id));
                    if (term != null)
                    {
                        term.CustomSortOrder = count;
                        count++;
                    }
                }
                termsToReturn = termsToReturn.OrderBy(t => t.CustomSortOrder).ToList();
            }


            return termsToReturn;
        }


        public override bool WillProvision(Web web, Model.ProvisioningTemplate template)
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

    internal class TermGroupsContext
    {
        public TermStore TermStore { get; set; }
        public TokenParser Parser { get; set; }
        public PnPMonitoredScope Scope { get; set; }
    }
}
