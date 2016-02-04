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

                foreach (var modelTermGroup in template.TermGroups)
                {
                    #region Group

                    var newGroup = false;

                    TermGroup group = termStore.Groups.FirstOrDefault(g => g.Id == modelTermGroup.Id);
                    if (group == null)
                    {
                        if (modelTermGroup.Name == "Site Collection")
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
                            var parsedGroupName = parser.ParseString(modelTermGroup.Name);
                            group = termStore.Groups.FirstOrDefault(g => g.Name == parsedGroupName);

                            if (group == null)
                            {
                                if (modelTermGroup.Id == Guid.Empty)
                                {
                                    modelTermGroup.Id = Guid.NewGuid();
                                }
                                group = termStore.CreateGroup(parsedGroupName, modelTermGroup.Id);

                                group.Description = modelTermGroup.Description;

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
                        if (!newGroup)
                        {
                            set = group.TermSets.FirstOrDefault(ts => ts.Id == modelTermSet.Id);
                            if (set == null)
                            {
                                set = group.TermSets.FirstOrDefault(ts => ts.Name == modelTermSet.Name);

                            }
                        }
                        if (set == null)
                        {
                            if (modelTermSet.Id == Guid.Empty)
                            {
                                modelTermSet.Id = Guid.NewGuid();
                            }
                            set = group.CreateTermSet(parser.ParseString(modelTermSet.Name), modelTermSet.Id, modelTermSet.Language ?? termStore.DefaultLanguage);
                            parser.AddToken(new TermSetIdToken(web, modelTermGroup.Name, modelTermSet.Name, modelTermSet.Id));
                            newTermSet = true;
                            set.IsOpenForTermCreation = modelTermSet.IsOpenForTermCreation;
                            set.IsAvailableForTagging = modelTermSet.IsAvailableForTagging;
                            foreach (var property in modelTermSet.Properties)
                            {
                                set.SetCustomProperty(property.Key, property.Value);
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
                                        term = terms.FirstOrDefault(t => t.Name == modelTerm.Name);
                                        if (term == null)
                                        {
                                            var returnTuple = CreateTerm<TermSet>(web, modelTerm, set, termStore, parser, scope);
                                            modelTerm.Id = returnTuple.Item1;
                                            parser = returnTuple.Item2;
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
                                }
                                else
                                {
                                    var returnTuple = CreateTerm<TermSet>(web, modelTerm, set, termStore, parser, scope);
                                    modelTerm.Id = returnTuple.Item1;
                                    parser = returnTuple.Item2;
                                }
                            }
                            else
                            {
                                var returnTuple = CreateTerm<TermSet>(web, modelTerm, set, termStore, parser, scope);
                                modelTerm.Id = returnTuple.Item1;
                                parser = returnTuple.Item2;
                            }
                        }

                        // do we need custom sorting?
                        if (modelTermSet.Terms.Any(t => t.CustomSortOrder > -1))
                        {
                            var sortedTerms = modelTermSet.Terms.OrderBy(t => t.CustomSortOrder);

                            var customSortString = sortedTerms.Aggregate(string.Empty, (a, i) => a + i.Id.ToString() + ":");
                            customSortString = customSortString.TrimEnd(new[] { ':' });

                            set.CustomSortOrder = customSortString;
                            termStore.CommitAll();
                            web.Context.ExecuteQueryRetry();
                        }
                    }

                    #endregion

                }
            }
            return parser;
        }

        private Tuple<Guid, TokenParser> CreateTerm<T>(Web web, Model.Term modelTerm, TaxonomyItem parent, TermStore termStore, TokenParser parser, PnPMonitoredScope scope) where T : TaxonomyItem
        {
            Term term;
            if (modelTerm.Id == Guid.Empty)
            {
                modelTerm.Id = Guid.NewGuid();
            }

            if (parent is Term)
            {
                var languages = termStore.DefaultLanguage;
                term = ((Term)parent).CreateTerm(parser.ParseString(modelTerm.Name), modelTerm.Language ?? termStore.DefaultLanguage, modelTerm.Id);

            }
            else
            {
                term = ((TermSet)parent).CreateTerm(parser.ParseString(modelTerm.Name), modelTerm.Language ?? termStore.DefaultLanguage, modelTerm.Id);
            }
            if (!String.IsNullOrEmpty(modelTerm.Description))
            {
                term.SetDescription(modelTerm.Description, modelTerm.Language ?? termStore.DefaultLanguage);
            }
            if (!String.IsNullOrEmpty(modelTerm.Owner))
            {
                term.Owner = modelTerm.Owner;
            }

            term.IsAvailableForTagging = modelTerm.IsAvailableForTagging;

            if (modelTerm.Properties.Any() || modelTerm.Labels.Any() || modelTerm.LocalProperties.Any())
            {
                if (modelTerm.Labels.Any())
                {
                    foreach (var label in modelTerm.Labels)
                    {
                        if ((label.IsDefaultForLanguage && label.Language != termStore.DefaultLanguage) || label.IsDefaultForLanguage == false)
                        {
                            var l = term.CreateLabel(parser.ParseString(label.Value), label.Language, label.IsDefaultForLanguage);
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language);
                            WriteWarning(string.Format(CoreResources.Provisioning_ObjectHandlers_TermGroups_Skipping_label__0___label_is_to_set_to_default_for_language__1__while_the_default_termstore_language_is_also__1_, label.Value, label.Language), ProvisioningMessageType.Warning);
                        }
                    }
                }

                if (modelTerm.Properties.Any())
                {
                    foreach (var property in modelTerm.Properties)
                    {
                        term.SetCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
                    }
                }
                if (modelTerm.LocalProperties.Any())
                {
                    foreach (var property in modelTerm.LocalProperties)
                    {
                        term.SetLocalCustomProperty(parser.ParseString(property.Key), parser.ParseString(property.Value));
                    }
                }
            }
            termStore.CommitAll();

            web.Context.Load(term);
            web.Context.ExecuteQueryRetry();

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
                                modelTermTerm.Id = returnTuple.Item1;
                                parser = returnTuple.Item2;
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
                        modelTermTerm.Id = returnTuple.Item1;
                        parser = returnTuple.Item2;
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
            return Tuple.Create(modelTerm.Id, parser);
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
					web.Context.Load(termStore, t => t.Id, t => t.DefaultLanguage);
					web.Context.ExecuteQueryRetry();

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
}
