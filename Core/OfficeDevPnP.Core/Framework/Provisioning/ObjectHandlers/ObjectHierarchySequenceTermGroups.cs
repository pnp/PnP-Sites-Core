#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using Term = Microsoft.SharePoint.Client.Taxonomy.Term;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectHierarchySequenceTermGroups : ObjectHierarchyHandlerBase
    {
        private List<TermGroupHelper.ReusedTerm> reusedTerms;

        public override string Name => "Term Groups";

        public override TokenParser ProvisionObjects(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser,
            ApplyConfiguration configuration)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                foreach (var sequence in hierarchy.Sequences)
                {

                    this.reusedTerms = new List<TermGroupHelper.ReusedTerm>();

                    var context = tenant.Context as ClientContext;
                    
                    TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(context);
                    TermStore termStore = null;

                    try
                    {
                        termStore = taxSession.GetDefaultKeywordsTermStore();
                        context.Load(termStore,
                            ts => ts.Languages,
                            ts => ts.DefaultLanguage,
                            ts => ts.Groups.Include(
                                tg => tg.Name,
                                tg => tg.Id,
                                tg => tg.TermSets.Include(
                                    tset => tset.Name,
                                    tset => tset.Id)));
                        context.ExecuteQueryRetry();
                    }
                    catch (ServerException)
                    {
                        WriteMessage("Failed to retrieve existing terms", ProvisioningMessageType.Error);

                        return parser;
                    }

                    foreach (var modelTermGroup in sequence.TermStore.TermGroups)
                    {
                        this.reusedTerms.AddRange(TermGroupHelper.ProcessGroup(context, taxSession, termStore, modelTermGroup, null, parser, scope));
                    }

                    foreach (var reusedTerm in this.reusedTerms)
                    {
                        TermGroupHelper.TryReuseTerm(context, reusedTerm.ModelTerm, reusedTerm.Parent, reusedTerm.TermStore, parser, scope);
                    }
                }
            }
            return parser;
        }

        private class TryReuseTermResult
        {
            public bool Success { get; set; }
            public TokenParser UpdatedParser { get; set; }
        }

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ExtractConfiguration configuration)
        {
            throw new NotImplementedException();
        }

        private List<Model.Term> GetTerms<T>(ClientRuntimeContext context, TaxonomyItem parent, int defaultLanguage, Boolean isSiteCollectionTermGroup = false)
        {
            List<Model.Term> termsToReturn = new List<Model.Term>();
            Microsoft.SharePoint.Client.Taxonomy.TermCollection terms;
            var customSortOrder = string.Empty;
            if (parent is Microsoft.SharePoint.Client.Taxonomy.TermSet)
            {
                terms = ((Microsoft.SharePoint.Client.Taxonomy.TermSet)parent).Terms;
                customSortOrder = ((Microsoft.SharePoint.Client.Taxonomy.TermSet)parent).CustomSortOrder;
            }
            else
            {
                terms = ((Microsoft.SharePoint.Client.Taxonomy.Term)parent).Terms;
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


        public override bool WillProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration)
        {
            return hierarchy.Sequences.Where(s => s.TermStore != null && s.TermStore.TermGroups != null && s.TermStore.TermGroups.Any()).Any();
        }

        public override bool WillExtract(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ExtractConfiguration configuration)
        {
            return false;
        }
    }
}
#endif