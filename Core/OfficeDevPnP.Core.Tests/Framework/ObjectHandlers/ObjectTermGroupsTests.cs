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

        [TestInitialize]
        public void Initialize()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                _termSetGuid = Guid.Parse("355f020a-c4a4-43da-a314-986826bddc38");
                _termGroupGuid = Guid.Parse("f3c879f2-c065-4aca-9a7d-b89ffc3c47f9");
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

                var store = session.GetDefaultKeywordsTermStore();

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

            var term1 = new Term(Guid.Parse("a091856e-761a-4123-b0bf-f734168e2be9"), "TestProvisioningTerm 1", null, null, null, null, null);
            term1.Properties.Add("TestProp1", "Test Value 1");
            term1.LocalProperties.Add("TestLocalProp1", "Test Value 1");
            term1.Labels.Add(new TermLabel() { Language = 1033, Value = "Testing" });

            term1.Terms.Add(new Term(Guid.Parse("c5828235-d524-4b9f-88a9-76c8e52c8fd0"), "Sub Term 1", null, null, null, null, null));

            terms.Add(term1);

            terms.Add(new Term(Guid.Parse("08a5083c-8cb9-4161-b47a-f67f81876b10"), "TestProvisioningTerm 2", null, null, null, null, null));

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
                ctx.Load(group, g => g.Name, g => g.Id, g => g.TermSets.Include(
                        tset => tset.Name,
                        tset => tset.Id));
                ctx.ExecuteQuery();
                var set = group.TermSets.FirstOrDefault(ts => ts.Id == termSet.Id || ts.Name == termSet.Name);
                ctx.Load(set.Terms, s=>s.Include(y=>y.Terms, y=>y.Id, y=>y.Name));
                ctx.ExecuteQuery();
                Assert.IsInstanceOfType(group, typeof(Microsoft.SharePoint.Client.Taxonomy.TermGroup));
                Assert.IsInstanceOfType(set, typeof(Microsoft.SharePoint.Client.Taxonomy.TermSet));
                Assert.IsTrue(set.Terms[0].Id == term1.Id);
                Assert.IsTrue(set.Terms[0].Terms[0].Id == term1.Terms[0].Id);
                Assert.IsTrue(set.Terms[1].Name == terms[1].Name);

                term1.Terms.Add(new Term(Guid.Parse("e9560121-e53d-4881-862b-f82362e79090"), "Sub Term 2", null, null, null, null, null) );
                new ObjectTermGroups().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                ctx.Load(set.Terms, s => s.Include(y => y.Terms, y => y.Id, y => y.Name));
                ctx.ExecuteQuery();

                Assert.IsTrue(set.Terms[0].Id == term1.Id);
                Assert.IsTrue(set.Terms[0].Terms[0].Id == term1.Terms[0].Id);
                Assert.IsTrue(set.Terms[0].Terms[1].Id == term1.Terms[1].Id);
                Assert.IsTrue(set.Terms[1].Name == terms[1].Name);

                var template2 = new ProvisioningTemplate();
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };
                template2 = new ObjectTermGroups().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.TermGroups.Any());
                Assert.IsInstanceOfType(template.TermGroups, typeof(Core.Framework.Provisioning.Model.TermGroupCollection));
            }


        }

    }
}
