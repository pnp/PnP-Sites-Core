using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Taxonomy tests are not supported when testing using app-only");
                return;
            }

            _termSetGuid = Guid.NewGuid();
            _termGroupGuid = Guid.NewGuid();
            _additionalTermSetGuid = Guid.NewGuid();
        }

        [TestCleanup]
        public void CleanUp()
        {
            if (TestCommon.AppOnlyTesting())
            {
                return;
            }

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
                }
                catch
                {
                }

                if (_termGroupGuid != Guid.Empty)
                {
                    try
                    {
                        TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
                        var store = session.GetDefaultSiteCollectionTermStore();

                        var termGroup = store.GetGroup(_termGroupGuid);
                        termGroup.DeleteObject();

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

            var termGroup = new TermGroup(_termGroupGuid, "{sitecollectiontermgroupname}", null);

            var termSets = new List<TermSet>();

            var termSet = new TermSet(_termSetGuid, "TestProvisioningTermSet", null, true, false, null, null);
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

            var termGroup = new TermGroup(_termGroupGuid, "TestProvisioningGroup", null);

            var termSets = new List<TermSet>();

            var termSet = new TermSet(_termSetGuid, "TestProvisioningTermSet", null, true, false, null, null);

            var terms = new List<Term>();

            var term1Id = Guid.NewGuid();
            const string term1Name = "TestProvisioningTerm 1";
            var term1 = new Term(term1Id, term1Name, null, null, null, null, null);
            term1.Properties.Add("TestProp1", "Test Value 1");
            term1.LocalProperties.Add("TestLocalProp1", "Test Value 1");
            term1.Labels.Add(new TermLabel() { Language = 1033, Value = "Testing" });

            var term1Subterm1Id = Guid.NewGuid();
            term1.Terms.Add(new Term(term1Subterm1Id, "Sub Term 1", null, null, null, null, null));

            terms.Add(term1);

            var term2Id = Guid.NewGuid();
            const string term2Name = "TestProvisioningTerm 2";
            terms.Add(new Term(term2Id, term2Name, null, null, null, null, null));

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

                var orderedTerms = set.Terms.OrderBy(t => t.Name, StringComparer.Ordinal).ToArray();
                Assert.AreEqual(2, orderedTerms.Length);

                var remoteTerm1 = orderedTerms[0];
                Assert.AreEqual(term1Id, remoteTerm1.Id);
                StringAssert.Matches(remoteTerm1.Name, new Regex(Regex.Escape(term1Name)));

                var remoteTerm2 = orderedTerms[1];
                Assert.AreEqual(term2Id, remoteTerm2.Id);
                StringAssert.Matches(remoteTerm2.Name, new Regex(Regex.Escape(term2Name)));

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

            var termGroup = new TermGroup(_termGroupGuid, "TestProvisioningGroup", null);

            var termSets = new List<TermSet>();

            var termSet1 = new TermSet(_termSetGuid, "TestProvisioningTermSet1", null, true, false, null, null);
            var termSet2 = new TermSet(_additionalTermSetGuid, "TestProvisioningTermSet2", null, true, false, null, null);

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

            // check result by reading the template again
            using (var ctx = TestCommon.CreateClientContext())
            {
                var result = ctx.Web.GetProvisioningTemplate(new ProvisioningTemplateCreationInformation(ctx.Web)
                {
                    HandlersToProcess = Handlers.TermGroups,
                    IncludeAllTermGroups = true // without this being true no term groups will be returned
                });

                // note: cannot use TermGroupValidator class to validate the result as XML since the read template contains additional information like Description="", Owner="[...]", differing TermGroup ID etc. which makes the validation fail; so manually compare what's interesting
                var newTermGroups = result.TermGroups.Where(tg => tg.Name == termGroup.Name);
                Assert.AreEqual(1, newTermGroups.Count());

                var newTermGroup = newTermGroups.First();
                Assert.AreEqual(2, newTermGroup.TermSets.Count);
                Assert.AreEqual(1, newTermGroup.TermSets[0].Terms.Count);
                Assert.AreEqual(1, newTermGroup.TermSets[1].Terms.Count);

                // note: this check that the IDs of the source and reused term are the same to document this behavior
                Assert.AreEqual(sourceTerm.Id, newTermGroup.TermSets[0].Terms[0].Id);
                Assert.AreEqual(sourceTerm.Id, newTermGroup.TermSets[1].Terms[0].Id);
                Assert.IsTrue(newTermGroup.TermSets[0].Terms[0].IsReused);
                Assert.IsTrue(newTermGroup.TermSets[1].Terms[0].IsReused);
            }
        }

        [TestMethod]
        public void CanProvisionTokenizedTermsTwiceIdMatch()
        {
            var template = new ProvisioningTemplate();
            const string termGroupName = "TestProvisioningGroup";
            var termGroup = new TermGroup(_termGroupGuid, termGroupName, null);

            const string termSiteName = "TestProvisioningTermSet - {sitename}";
            var termSet1 = new TermSet(_termSetGuid, termSiteName, null, true, false, null, null);

            var term1Id = Guid.NewGuid();
            const string term1Name = "TestProvisioningTerm - {siteid}";
            var term1 = new Term(term1Id, term1Name, null, null, null, null, null);

            termSet1.Terms.Add(term1);
            termGroup.TermSets.Add(termSet1);
            template.TermGroups.Add(termGroup);

            for (int index = 0; index < 2; index++)
            {
                using (ClientContext ctx = TestCommon.CreateClientContext())
                {
                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                    TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                    var store = session.GetDefaultKeywordsTermStore();
                    var group = store.GetGroup(_termGroupGuid);
                    var set = store.GetTermSet(_termSetGuid);

                    ctx.Load(group);
                    ctx.Load(set, s => s.Id, s => s.Name, s => s.Terms);
                    ctx.ExecuteQueryRetry();

                    Assert.IsInstanceOfType(group, typeof(Microsoft.SharePoint.Client.Taxonomy.TermGroup));
                    StringAssert.Matches(group.Name, new Regex(Regex.Escape(termGroupName)));
                    Assert.AreEqual(_termGroupGuid, group.Id);

                    Assert.IsInstanceOfType(set, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                    Assert.AreEqual(1, set.Terms.Count);
                    Assert.AreEqual(_termSetGuid, set.Id);
                    StringAssert.DoesNotMatch(set.Name, new Regex(Regex.Escape(termSiteName)));

                    var remoteTerm1 = set.Terms[0];
                    Assert.AreEqual(term1Id, remoteTerm1.Id);
                    StringAssert.DoesNotMatch(remoteTerm1.Name, new Regex(Regex.Escape(term1Name)));
                }
            }
        }

        [TestMethod]
        public void CanProvisionTokenizedTermsTwiceNameMatch()
        {
            var template = new ProvisioningTemplate();
            const string termGroupName = "TestProvisioningGroup";
            var termGroup = new TermGroup(_termGroupGuid, termGroupName, null);

            const string termSiteName = "TestProvisioningTermSet - {sitename}";
            var termSet1 = new TermSet(_termSetGuid, termSiteName, null, true, false, null, null);

            var term1Id = Guid.NewGuid();
            const string term1Name = "TestProvisioningTerm - {siteid}";
            var term1 = new Term(term1Id, term1Name, null, null, null, null, null);

            termSet1.Terms.Add(term1);
            termGroup.TermSets.Add(termSet1);
            template.TermGroups.Add(termGroup);

            for (int index = 0; index < 2; index++)
            {
                if (index == 1)
                {
                    // Assign a new ID to the Term to test the name matching logic.
                    term1.Id = Guid.NewGuid();
                }

                using (ClientContext ctx = TestCommon.CreateClientContext())
                {
                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                    TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                    var store = session.GetDefaultKeywordsTermStore();
                    var group = store.GetGroup(_termGroupGuid);
                    var set = store.GetTermSet(_termSetGuid);

                    ctx.Load(group);
                    ctx.Load(set, s => s.Id, s => s.Name, s => s.Terms);
                    ctx.ExecuteQueryRetry();

                    Assert.IsInstanceOfType(group, typeof(Microsoft.SharePoint.Client.Taxonomy.TermGroup));
                    StringAssert.Matches(group.Name, new Regex(Regex.Escape(termGroupName)));
                    Assert.AreEqual(_termGroupGuid, group.Id);

                    Assert.IsInstanceOfType(set, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                    Assert.AreEqual(1, set.Terms.Count);
                    Assert.AreEqual(_termSetGuid, set.Id);
                    StringAssert.DoesNotMatch(set.Name, new Regex(Regex.Escape(termSiteName)));

                    var remoteTerm1 = set.Terms[0];
                    Assert.AreEqual(term1Id, remoteTerm1.Id);
                    StringAssert.DoesNotMatch(remoteTerm1.Name, new Regex(Regex.Escape(term1Name)));
                }
            }
        }
    }
}
