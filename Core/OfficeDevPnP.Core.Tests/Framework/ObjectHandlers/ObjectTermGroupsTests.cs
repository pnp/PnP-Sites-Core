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
        private Guid _termGroupGuid;
        private Guid _termSetReusable1Guid;
        private Guid _termSetReusable2Guid;
        private Guid _termGroupReusableGuid;

        [TestInitialize]
        public void Initialize()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                _termSetGuid = Guid.NewGuid();
                _termGroupGuid = Guid.NewGuid();
                _termSetReusable1Guid = Guid.NewGuid();
                _termSetReusable2Guid = Guid.NewGuid();
                _termGroupReusableGuid = Guid.NewGuid();
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
                        var termSet = store.GetTermSet(_termSetGuid);
                        termSet.DeleteObject();

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

            var term1 = new Term(Guid.NewGuid(), "TestProvisioningTerm 1", null, null, null, null, null, false, Guid.Empty);
            term1.Properties.Add("TestProp1", "Test Value 1");
            term1.LocalProperties.Add("TestLocalProp1", "Test Value 1");
            term1.Labels.Add(new TermLabel() { Language = 1033, Value = "Testing" });

            term1.Terms.Add(new Term(Guid.NewGuid(), "Sub Term 1", null, null, null, null, null, false, Guid.Empty));

            terms.Add(term1);

            terms.Add(new Term(Guid.NewGuid(), "TestProvisioningTerm 2", null, null, null, null, null, false, Guid.Empty));

            termSet.Terms.AddRange(terms);

            termSets.Add(termSet);

            termGroup.TermSets.AddRange(termSets);

            template.TermGroups.Add(termGroup);

            using (var ctx = TestCommon.CreateClientContext())
            {

                var parser = new TokenParser(ctx.Web, template);

                new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                var store = session.GetDefaultSiteCollectionTermStore();

                var group = store.GetGroup(_termGroupGuid);
                var set = store.GetTermSet(_termSetGuid);

                ctx.Load(group);
                ctx.Load(set, ts => ts.Terms);

                ctx.ExecuteQueryRetry();

                Assert.IsInstanceOfType(group, typeof(Microsoft.SharePoint.Client.Taxonomy.TermGroup));
                Assert.IsInstanceOfType(set, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                Assert.IsTrue(set.Terms.Count == 2);


                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template2 = new ProvisioningTemplate();
                template2 = new ObjectTermGroups().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.TermGroups.Any());
                Assert.IsInstanceOfType(template.TermGroups, typeof(List<TermGroup>));
            }


        }

        [TestMethod]
        public void CanProvisionResuableTerms()
        {
            var template = new ProvisioningTemplate();

            TermGroup termGroup = new TermGroup(_termGroupReusableGuid, "TestProvisioningGroupReusable", null);

            List<TermSet> termSets = new List<TermSet>();

            TermSet termSet1 = new TermSet(_termSetReusable1Guid, "TestProvisioningTermSetReusable1", null, true, false, null, null);
            TermSet termSet2 = new TermSet(_termSetReusable2Guid, "TestProvisioningTermSetReusable2", null, true, false, null, null);

            List<Term> terms1 = new List<Term>();
            List<Term> terms2 = new List<Term>();

            Guid reusedTermGuid = Guid.NewGuid();

            var term1 = new Term(reusedTermGuid, "TestProvisioningReusedTerm 1", null, null, null, null, null, true, reusedTermGuid);
            term1.Properties.Add("TestProp1", "Test Value 1");
            term1.LocalProperties.Add("TestLocalProp1", "Test Value 1");
            term1.Labels.Add(new TermLabel() { Language = 1033, Value = "Testing" });

            terms1.Add(term1);
            termSet1.Terms.AddRange(terms1);
            termSets.Add(termSet1);

            var term2 = new Term(reusedTermGuid, "TestProvisioningReusedTerm 1", null, null, null, null, null, true, reusedTermGuid);
            term1.Properties.Add("TestProp1", "Test Value 1");
            term1.LocalProperties.Add("TestLocalProp1", "Test Value 1");
            term1.Labels.Add(new TermLabel() { Language = 1033, Value = "Testing" });

            terms2.Add(term1);
            termSet2.Terms.AddRange(terms2);
            termSets.Add(termSet2);

            termGroup.TermSets.AddRange(termSets);
            template.TermGroups.Add(termGroup);

            using (var ctx = TestCommon.CreateClientContext())
            {

                var parser = new TokenParser(ctx.Web, template);

                new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);

                var store = session.GetDefaultSiteCollectionTermStore();

                var group = store.GetGroup(_termGroupReusableGuid);
                var set1 = store.GetTermSet(_termSetReusable1Guid);
                var set2 = store.GetTermSet(_termSetReusable2Guid);

                ctx.Load(group);
                ctx.Load(set1, ts => ts.Terms);
                ctx.Load(set2, ts => ts.Terms);

                ctx.ExecuteQueryRetry();

                Assert.IsInstanceOfType(group, typeof(Microsoft.SharePoint.Client.Taxonomy.TermGroup));
                Assert.IsInstanceOfType(set1, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                Assert.IsTrue(set1.Terms.Count == 1);
                Assert.IsInstanceOfType(set2, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                Assert.IsTrue(set2.Terms.Count == 1);
            }


        }

    }
}
