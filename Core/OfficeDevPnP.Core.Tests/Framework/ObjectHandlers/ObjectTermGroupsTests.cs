using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using Term = OfficeDevPnP.Core.Framework.Provisioning.Model.Term;
using TermGroup = OfficeDevPnP.Core.Framework.Provisioning.Model.TermGroup;
using TermSet = OfficeDevPnP.Core.Framework.Provisioning.Model.TermSet;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectTermGroupsTests
    {

        private Guid _termSetGuid;
        private Guid _additionalTermSetGuid;
        private Guid _termGroupGuid;

        [TestInitialize]
        public void Initialize()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                _termSetGuid = Guid.NewGuid();
                _termGroupGuid = Guid.NewGuid();
                _additionalTermSetGuid = Guid.NewGuid();
            }
            else
            {
                Assert.Inconclusive("Taxonomy tests are not supported when testing using app-only");
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    try
                    {
                        TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                        var store = session.GetDefaultSiteCollectionTermStore();
                        var termSet1 = store.GetTermSet(_termSetGuid);
                        var termSet2 = store.GetTermSet(_additionalTermSetGuid);

                        termSet1.DeleteObject();
                        termSet2.DeleteObject();

                        store.CommitAll();
                        ctx.ExecuteQueryRetry();

                        if (_termGroupGuid != Guid.Empty)
                        {
                            var termGroup = store.GetGroup(_termGroupGuid);
                            termGroup.DeleteObject();
                        }
                        store.CommitAll();
                        ctx.ExecuteQueryRetry();
                    }
                    catch
                    {
                    }
                }
            }
        }

        [TestMethod]
        public void CanProvisionToSiteCollectionTermGroupUsingToken()
        {
            var template = new ProvisioningTemplate();
            _termGroupGuid = Guid.Empty;

            TermGroup termGroup = new TermGroup(_termGroupGuid, "{sitecollectiontermgroupname}", null);

            List<TermSet> termSets = new List<TermSet>();

            TermSet termSet = new TermSet(_termSetGuid, "TestProvisioningTermSet", null, true, false, null, null);
            termSets.Add(termSet);

            termGroup.TermSets.AddRange(termSets);

            template.TermGroups.Add(termGroup);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                var store = session.GetDefaultSiteCollectionTermStore();

                var set = store.GetTermSet(_termSetGuid);
                var group = set.Group;

                ctx.Load(set);
                ctx.Load(group);
                ctx.ExecuteQueryRetry();

                Assert.IsInstanceOfType(group, typeof(Microsoft.SharePoint.Client.Taxonomy.TermGroup));
                Assert.IsInstanceOfType(set, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                Assert.IsTrue(group.IsSiteCollectionGroup);

            }

        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();

            TermGroup termGroup = new TermGroup(_termGroupGuid, "TestProvisioningGroup", null);

            List<TermSet> termSets = new List<TermSet>();

            TermSet termSet = new TermSet(_termSetGuid, "TestProvisioningTermSet", null, true, false, null, null);

            List<Term> terms = new List<Term>();

            var term1 = new Term(Guid.NewGuid(), "TestProvisioningTerm 1", null, null, null, null, null);
            term1.Properties.Add("TestProp1", "Test Value 1");
            term1.LocalProperties.Add("TestLocalProp1", "Test Value 1");
            term1.Labels.Add(new TermLabel() { Language = 1033, Value = "Testing" });

            term1.Terms.Add(new Term(Guid.NewGuid(), "Sub Term 1", null, null, null, null, null));

            terms.Add(term1);

            terms.Add(new Term(Guid.NewGuid(), "TestProvisioningTerm 2", null, null, null, null, null));

            termSet.Terms.AddRange(terms);

            termSets.Add(termSet);

            termGroup.TermSets.AddRange(termSets);

            template.TermGroups.Add(termGroup);

            using (var ctx = TestCommon.CreateClientContext())
            {

                var parser = new TokenParser(ctx.Web, template);

                new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                var store = session.GetDefaultKeywordsTermStore(); 
                var group = store.GetGroup(_termGroupGuid);
                var set = store.GetTermSet(_termSetGuid);

                ctx.Load(group);
                ctx.Load(set, s => s.Terms);
                ctx.ExecuteQueryRetry();

                Assert.IsInstanceOfType(group, typeof(Microsoft.SharePoint.Client.Taxonomy.TermGroup));
                Assert.IsInstanceOfType(set, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                Assert.IsTrue(set.Terms.Count == 2);


                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template2 = new ProvisioningTemplate();
                template2 = new ObjectTermGroups().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.TermGroups.Any());
                Assert.IsInstanceOfType(template.TermGroups, typeof(Core.Framework.Provisioning.Model.TermGroupCollection));
            }


        }

        [TestMethod]
        public void CanProvisionReusableTerms()
        {
            var template = new ProvisioningTemplate();

            TermGroup termGroup = new TermGroup(_termGroupGuid, "TestProvisioningGroup", null);

            List<TermSet> termSets = new List<TermSet>();

            TermSet termSet1 = new TermSet(_termSetGuid, "TestProvisioningTermSet1", null, true, false, null, null);
            TermSet termSet2 = new TermSet(_additionalTermSetGuid, "TestProvisioningTermSet2", null, true, false, null, null);

            var sourceTerm = new Term(Guid.NewGuid(), "Source Term 1", null, null, null, null, null)
            {
                IsReused = true,
                IsSourceTerm = true
            };


            var reusedTerm = new Term(sourceTerm.Id, "Source Term 1", null, null, null, null, null)
            {
                IsReused = true,
                SourceTermId = sourceTerm.Id
            };


            termSet1.Terms.Add(reusedTerm);
            termSet2.Terms.Add(sourceTerm);

            termSets.Add(termSet1);
            termSets.Add(termSet2);

            termGroup.TermSets.AddRange(termSets);

            template.TermGroups.Add(termGroup);

            using (var ctx = TestCommon.CreateClientContext())
            {

                var parser = new TokenParser(ctx.Web, template);

                new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                var store = session.GetDefaultKeywordsTermStore();
                var group = store.GetGroup(_termGroupGuid);
                var set2 = store.GetTermSet(_additionalTermSetGuid);

                ctx.Load(group);
                ctx.Load(set2, s => s.Terms);
                ctx.ExecuteQueryRetry();

                var createdSourceTerm = set2.GetTerm(sourceTerm.Id);
                ctx.Load(createdSourceTerm);
                ctx.ExecuteQueryRetry();

                Assert.IsTrue(createdSourceTerm.IsSourceTerm);
                Assert.IsTrue(createdSourceTerm.IsReused);

                var set1 = store.GetTermSet(_termSetGuid);
                ctx.Load(set1, s => s.Terms);
                ctx.ExecuteQueryRetry();

                var createdReusedTerm = set1.GetTerm(reusedTerm.Id);
                ctx.Load(createdReusedTerm, c => c.SourceTerm, c => c.IsReused);
                ctx.ExecuteQueryRetry();
                Assert.IsTrue(createdReusedTerm.SourceTerm.Id == sourceTerm.Id);
                Assert.IsTrue(createdReusedTerm.IsReused);
            }


        }

    }
}
