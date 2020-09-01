using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Tests;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Linq;
using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;

namespace Microsoft.SharePoint.Client.Tests
{
    [TestClass()]
    public class TaxonomyExtensionsTests
    {
        private string _termGroupName; // For easy reference. Set in the Initialize method
        private string _termSetName1; // For easy reference. Set in the Initialize method
        private string _termSetName2; // For easy reference. Set in the Initialize method
        private string _termName; // For easy reference. Set in the Initialize method
        private Guid _termGroupId = Guid.NewGuid(); //Hardcoded GUIDs had sideffects when running tests. Several successive taxonomy tests would trigger a server exception.
        private Guid _termSet1Id = Guid.NewGuid(); //Hardcoded GUIDs had sideffects when running tests. Several successive taxonomy tests would trigger a server exception.
        private Guid _termSet2Id = Guid.NewGuid();
        private Guid _termId = Guid.NewGuid(); //Hardcoded GUIDs had sideffects when running tests. Several successive taxonomy tests would trigger a server exception and term field value label was untestable as previous term label in hidden field would overwrite new term label in list item.

        private Guid _listId; // For easy reference

        private string SampleTermSetPath = "../../Resources/ImportTermSet.csv";
        private string SampleUpdateTermSetPath = "../../Resources/UpdateTermSet.csv";
        private string SampleGuidTermSetPath = "../../Resources/GuidTermSet.csv";
        private Guid UpdateTermSetId = new Guid("{35585956-83E4-4A44-8FC5-AC50942E3187}");
        private Guid GuidTermSetId = new Guid("{90FD4208-8281-40CC-872E-DD85F33B50AB}");

        #region Test initialize and cleanup
        [TestInitialize]
        public void Initialize()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                Console.WriteLine("TaxonomyExtensionsTests.Initialise");
                // Create some taxonomy groups and terms
                using (var clientContext = TestCommon.CreateClientContext())
                {
                    _termGroupName = "Test_Group_" + DateTime.Now.ToFileTime();
                    _termSetName1 = "Test_Termset_1_" + DateTime.Now.ToFileTime();
                    _termSetName2 = "Test_Termset_2_" + DateTime.Now.ToFileTime();
                    _termName = "Test_Term_" + DateTime.Now.ToFileTime();

                    var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                    var termStore = taxSession.GetDefaultSiteCollectionTermStore();

                    // Termgroup
                    // Does the termgroup exist?
                    var termGroup = termStore.GetGroup(_termGroupId);
                    clientContext.Load(termGroup, g => g.Id);
                    clientContext.ExecuteQueryRetry();

                    // Create if non existant
                    if (termGroup.ServerObjectIsNull.Value)
                    {
                        termGroup = termStore.CreateGroup(_termGroupName, _termGroupId);
                        clientContext.Load(termGroup);
                        clientContext.ExecuteQueryRetry();
                    }

                    // Termset
                    // Does the termset exist?
                    var termSet1 = termStore.GetTermSet(_termSet1Id);
                    clientContext.Load(termSet1, ts => ts.Id);
                    clientContext.ExecuteQueryRetry();

                    // Create if non existant
                    if (termSet1.ServerObjectIsNull.Value)
                    {
                        termSet1 = termGroup.CreateTermSet(_termSetName1, _termSet1Id, 1033);
                        clientContext.Load(termSet1);
                        clientContext.ExecuteQueryRetry();
                    }

                    // Termset
                    // Does the termset exist?
                    var termSet2 = termStore.GetTermSet(_termSet2Id);
                    clientContext.Load(termSet2, ts => ts.Id);
                    clientContext.ExecuteQueryRetry();

                    // Create if non existant
                    if (termSet2.ServerObjectIsNull.Value)
                    {
                        termSet2 = termGroup.CreateTermSet(_termSetName2, _termSet2Id, 1033);
                        clientContext.Load(termSet2);
                        clientContext.ExecuteQueryRetry();
                    }

                    // Term
                    // Does the term exist?
                    var term = termStore.GetTerm(_termId);
                    clientContext.Load(term, t => t.Id);
                    clientContext.ExecuteQueryRetry();

                    // Create if non existant
                    if (term.ServerObjectIsNull.Value)
                    {
                        term = termSet1.CreateTerm(_termName, 1033, _termId);
                        clientContext.ExecuteQueryRetry();
                    }
                    else
                    {
                        var label = term.GetDefaultLabel(1033);
                        clientContext.ExecuteQueryRetry();
                        _termName = label.Value;
                    }

                    // List
                    ListCreationInformation listCI = new ListCreationInformation();
                    listCI.TemplateType = (int)ListTemplateType.GenericList;
                    listCI.Title = "Test_List_" + DateTime.Now.ToFileTime();
                    var list = clientContext.Web.Lists.Add(listCI);
                    clientContext.Load(list);
                    clientContext.ExecuteQueryRetry();
                    _listId = list.Id;

                }
            }
            else
            {
                Assert.Inconclusive("Taxonomy tests are not supported when testing using app-only");
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                Console.WriteLine("TaxonomyExtensionsTests.Cleanup");


                // Clean up Taxonomy
                try
                {                    
                    this.CleanupTaxonomy();
                }
                catch (ServerException serverEx)
                {
                    if (!string.IsNullOrEmpty(serverEx.ServerErrorTypeName)
                        && serverEx.ServerErrorTypeName.Contains("TermStoreErrorCodeEx"))
                    {
                        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(2));
                        this.CleanupTaxonomy();
                    }
                }

                using (var clientContext = TestCommon.CreateClientContext())
                {
                    // Clean up fields
                    var fields = clientContext.LoadQuery(clientContext.Web.Fields);
                    clientContext.ExecuteQueryRetry();
                    var testFields = fields.Where(f => f.InternalName.StartsWith("Test_", StringComparison.OrdinalIgnoreCase));
                    foreach (var field in testFields)
                    {
                        field.DeleteObject();
                    }
                    clientContext.ExecuteQueryRetry();

                    // Clean up list
                    var list = clientContext.Web.Lists.GetById(_listId);
                    list.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }

        private void CleanupTaxonomy()
        {
            if (!TestCommon.AppOnlyTesting())
            {
                // Ensure that the group is empty before deleting it. 
                // exceptions like the following happen:
                // Microsoft.SharePoint.Client.ServerException: Microsoft.SharePoint.Client.ServerException: A Group cannot be deleted unless it is empty..

                OfficeDevPnP.Core.Tests.Utilities.RetryHelper.Do(
                    () => this.InnerCleanupTaxonomy(),
                    TimeSpan.FromSeconds(30),
                    3);
            }
        }

        private void InnerCleanupTaxonomy()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Clean up Taxonomy
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                var termGroup = termStore.GetGroup(_termGroupId);
                var termSets = termGroup.TermSets;
                clientContext.Load(termSets);
                clientContext.ExecuteQueryRetry();

                foreach (var termSet in termSets)
                {
                    termSet.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }

                termStore.CommitAll();
                clientContext.ExecuteQueryRetry();

                // termStore.UpdateCache();
                taxSession.UpdateCache();

                termGroup.DeleteObject(); // Will delete underlying termset
                clientContext.ExecuteQueryRetry();
            }
        }
        #endregion

        #region Create taxonomy field tests
        [TestMethod()]
        public void CreateTaxonomyFieldTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve Termset
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                clientContext.Load(termSet);
                clientContext.ExecuteQueryRetry();

                // Get Test TermSet

                var web = clientContext.Web;
                var fieldName = "Test_" + DateTime.Now.ToFileTime();
                var fieldId = Guid.NewGuid();

                TaxonomyFieldCreationInformation fieldCI = new TaxonomyFieldCreationInformation()
                {
                    Id = fieldId,
                    DisplayName = fieldName,
                    InternalName = fieldName,
                    Group = "Test Fields Group",
                    TaxonomyItem = termSet
                };
                var field = web.CreateTaxonomyField(fieldCI);

                Assert.AreEqual(fieldId, field.Id, "Field IDs do not match.");
                Assert.AreEqual(fieldName, field.InternalName, "Field internal names do not match.");
                Assert.AreEqual("TaxonomyFieldType", field.TypeAsString, "Failed to create a TaxonomyField object.");
            }
        }

        [TestMethod()]
        public void CreateTaxonomyFieldMultiValueTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve Termset
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                clientContext.Load(termSet);
                clientContext.ExecuteQueryRetry();

                // Get Test TermSet

                var web = clientContext.Web;
                var fieldName = "Test_" + DateTime.Now.ToFileTime();
                var fieldId = Guid.NewGuid();
                TaxonomyFieldCreationInformation fieldCI = new TaxonomyFieldCreationInformation()
                {
                    Id = fieldId,
                    DisplayName = fieldName,
                    InternalName = fieldName,
                    Group = "Test Fields Group",
                    TaxonomyItem = termSet,
                    MultiValue = true
                };
                var field = web.CreateTaxonomyField(fieldCI);


                Assert.AreEqual(fieldId, field.Id, "Field IDs do not match.");
                Assert.AreEqual(fieldName, field.InternalName, "Field internal names do not match.");
                Assert.AreEqual("TaxonomyFieldTypeMulti", field.TypeAsString, "Failed to create a TaxonomyField object.");
            }
        }

        [TestMethod()]
        public void SetTaxonomyFieldValueTest()
        {
            var fieldName = "Test2_" + DateTime.Now.ToFileTime();

            var fieldId = Guid.NewGuid();

            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve list
                var list = clientContext.Web.Lists.GetById(_listId);
                clientContext.Load(list);
                clientContext.ExecuteQueryRetry();

                // Retrieve Termset
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                clientContext.Load(termSet);
                clientContext.ExecuteQueryRetry();

                // Create taxonomyfield first
                TaxonomyFieldCreationInformation fieldCI = new TaxonomyFieldCreationInformation()
                {
                    Id = fieldId,
                    DisplayName = fieldName,
                    InternalName = fieldName,
                    Group = "Test Fields Group",
                    TaxonomyItem = termSet
                };
                var field = list.CreateTaxonomyField(fieldCI);

                // Create Item
                ListItemCreationInformation itemCi = new ListItemCreationInformation();

                var item = list.AddItem(itemCi);
                item.Update();
                clientContext.Load(item);
                clientContext.ExecuteQueryRetry();

                item.SetTaxonomyFieldValue(fieldId, _termName, _termId);

                clientContext.Load(item, i => i[fieldName]);
                clientContext.ExecuteQueryRetry();

                var value = item[fieldName] as TaxonomyFieldValue;

                Assert.IsNotNull(value);
                Assert.IsTrue(value.WssId > 0, "Term WSS ID not set correctly");
                Assert.AreEqual(_termName, value.Label, "Term label not set correctly");
                Assert.AreEqual(_termId.ToString(), value.TermGuid, "Term GUID not set correctly");
            }
        }

        [TestMethod()]
        public void SetBlankTaxonomyFieldValueTest()
        {
            var fieldName = "Test2_" + DateTime.Now.ToFileTime();

            var fieldId = Guid.NewGuid();

            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve list
                var list = clientContext.Web.Lists.GetById(_listId);
                clientContext.Load(list);
                clientContext.ExecuteQueryRetry();

                // Retrieve Termset
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                clientContext.Load(termSet);
                clientContext.ExecuteQueryRetry();

                // Create taxonomyfield first
                TaxonomyFieldCreationInformation fieldCI = new TaxonomyFieldCreationInformation()
                {
                    Id = fieldId,
                    DisplayName = fieldName,
                    InternalName = fieldName,
                    Group = "Test Fields Group",
                    TaxonomyItem = termSet
                };
                //Add Enterprise keywords field so that we have at least two taxonomy fields in the list
                var keyworkdField = clientContext.Web.Fields.GetByInternalNameOrTitle("TaxKeyword");
                clientContext.Load(keyworkdField, f => f.Id);
                list.Fields.Add(keyworkdField);
                var field = list.CreateTaxonomyField(fieldCI);

                // Create Item
                ListItemCreationInformation itemCi = new ListItemCreationInformation();

                var item = list.AddItem(itemCi);
                item.Update();
                clientContext.Load(item);
                clientContext.ExecuteQueryRetry();

                //First set a valid value in at least two taxonomy fields (same term can be used if one field is the keyword field)
                item.SetTaxonomyFieldValue(fieldId, _termName, _termId);
                item.SetTaxonomyFieldValue(keyworkdField.Id, _termName, _termId);

                clientContext.Load(item, i => i[fieldName], i => i["TaxCatchAll"]);
                clientContext.ExecuteQueryRetry();

                Assert.AreEqual(2, (item["TaxCatchAll"] as FieldLookupValue[]).Length, "TaxCatchAll does not have 2 entries");
                var value = item[fieldName] as TaxonomyFieldValue;
                Assert.AreEqual(_termId.ToString(), value.TermGuid, "Term not set correctly");

                //Set a blank value in one of the taxonomy fields.
                item.SetTaxonomyFieldValue(fieldId, string.Empty, Guid.Empty);
                
                var taxonomyField = clientContext.CastTo<TaxonomyField>(field);
                clientContext.Load(taxonomyField, t => t.TextField);
                clientContext.Load(item, i => i[fieldName], i => i["TaxCatchAll"]);
                clientContext.ExecuteQueryRetry();

                var hiddenField = list.Fields.GetById(taxonomyField.TextField);
                clientContext.Load(hiddenField,
                    f => f.InternalName);
                clientContext.ExecuteQueryRetry();

                Assert.AreEqual(1, (item["TaxCatchAll"] as FieldLookupValue[]).Length, "TaxCatchAll does not have 1 entry");
                object taxonomyFieldValue = item[fieldName];
                object hiddenFieldValue = item[hiddenField.InternalName];
                Assert.IsNull(taxonomyFieldValue, "taxonomyFieldValue is not null");
                Assert.IsNull(hiddenFieldValue, "hiddenFieldValue is not null");
            }
        }

        [TestMethod()]
        public void SetTaxonomyFieldMultiValueTest()
        {
            var fieldName = "TaxKeyword";

            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve list
                var list = clientContext.Web.Lists.GetById(_listId);
                clientContext.Load(list);
                clientContext.ExecuteQueryRetry();

                // Retrieve Termset
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                clientContext.Load(termSet);
                clientContext.ExecuteQueryRetry();

                //Add Enterprise keywords field
                var keyworkdField = clientContext.Web.Fields.GetByInternalNameOrTitle(fieldName);
                clientContext.Load(keyworkdField, f => f.Id);
                list.Fields.Add(keyworkdField);

                //Create second term
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();

                Guid term2Id = Guid.NewGuid();
                string term2Name = "Test_Term_" + DateTime.Now.ToFileTime();

                var term2 = termStore.GetTerm(term2Id);
                clientContext.Load(term2, t => t.Id);
                clientContext.ExecuteQueryRetry();

                // Create if non existant
                if (term2.ServerObjectIsNull.Value)
                {
                    term2 = termSet.CreateTerm(term2Name, 1033, term2Id);
                    clientContext.ExecuteQueryRetry();
                }
                else
                {
                    var label2 = term2.GetDefaultLabel(1033);
                    clientContext.ExecuteQueryRetry();
                    term2Name = label2.Value;
                }

                // Create Item
                ListItemCreationInformation itemCi = new ListItemCreationInformation();

                var item = list.AddItem(itemCi);
                item.Update();
                clientContext.Load(item);
                clientContext.ExecuteQueryRetry();

                item.SetTaxonomyFieldValues(keyworkdField.Id, new List<KeyValuePair<Guid, string>> {
                    new KeyValuePair<Guid, string>(_termId, _termName),
                    new KeyValuePair<Guid, string>(term2Id, term2Name)
                });

                clientContext.Load(item, i => i[fieldName]);
                clientContext.ExecuteQueryRetry();

                var value = item[fieldName] as TaxonomyFieldValueCollection;

                Assert.IsNotNull(value);
                Assert.AreEqual(2, value.Count, "Taxonomy value count mismatch");
                Assert.IsTrue(value[0].WssId > 0, "Term WSS ID not set correctly");
                Assert.IsTrue(value[1].WssId > 0, "Term2 WSS ID not set correctly");
                Assert.AreEqual(_termName, value[0].Label, "Term label not set correctly");
                Assert.AreEqual(term2Name, value[1].Label, "Term2 label not set correctly");
                Assert.AreEqual(_termId.ToString(), value[0].TermGuid, "Term GUID not set correctly");
                Assert.AreEqual(term2Id.ToString(), value[1].TermGuid, "Term2 GUID not set correctly");
            }
        }

        [TestMethod()]
        public void SetBlankTaxonomyFieldMultiValueTest()
        {
            var fieldName = "TaxKeyword";

            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve list
                var list = clientContext.Web.Lists.GetById(_listId);
                clientContext.Load(list);
                clientContext.ExecuteQueryRetry();

                // Retrieve Termset
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                clientContext.Load(termSet);
                clientContext.ExecuteQueryRetry();

                //Add Enterprise keywords field
                var keywordField = clientContext.Web.Fields.GetByInternalNameOrTitle(fieldName);
                var keywordTaxonomyField = clientContext.CastTo<TaxonomyField>(keywordField);
                clientContext.Load(keywordTaxonomyField, f => f.Id, f => f.TextField);
                list.Fields.Add(keywordTaxonomyField);

                // Create Item
                ListItemCreationInformation itemCi = new ListItemCreationInformation();

                var item = list.AddItem(itemCi);
                item.Update();
                clientContext.Load(item);
                clientContext.ExecuteQueryRetry();

                item.SetTaxonomyFieldValues(keywordField.Id, new List<KeyValuePair<Guid, string>>());

                clientContext.Load(item, i => i[fieldName]);
                clientContext.ExecuteQueryRetry();

                var hiddenField = list.Fields.GetById(keywordTaxonomyField.TextField);
                clientContext.Load(hiddenField,
                    f => f.InternalName);
                clientContext.ExecuteQueryRetry();

                TaxonomyFieldValueCollection taxonomyFieldValueCollection = item[fieldName] as TaxonomyFieldValueCollection;
                object hiddenFieldValue = item[hiddenField.InternalName];
                Assert.AreEqual(0, taxonomyFieldValueCollection.Count, "taxonomyFieldValueCollection is not empty");
                Assert.IsNull(hiddenFieldValue, "hiddenFieldValue is not null");
            }
        }

        [TestMethod()]
        public void CreateTaxonomyFieldLinkedToTermSetTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve Termset
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                clientContext.Load(termSet);
                clientContext.ExecuteQueryRetry();

                var list = clientContext.Web.Lists.GetById(_listId);
                clientContext.Load(list);
                clientContext.ExecuteQueryRetry();

                var fieldName = "Test_" + DateTime.Now.ToFileTime();
                var fieldId = Guid.NewGuid();
                TaxonomyFieldCreationInformation fieldCI = new TaxonomyFieldCreationInformation()
                {
                    Id = fieldId,
                    DisplayName = fieldName,
                    InternalName = fieldName,
                    Group = "Test Fields Group",
                    TaxonomyItem = termSet
                };
                var field = list.CreateTaxonomyField(fieldCI);

                Assert.AreEqual(fieldId, field.Id, "Field IDs do not match.");
                Assert.AreEqual(fieldName, field.InternalName, "Field internal names do not match.");
                Assert.AreEqual("TaxonomyFieldType", field.TypeAsString, "Failed to create a TaxonomyField object.");
            }
        }

        [TestMethod()]
        public void CreateTaxonomyFieldLinkedToTermTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // Retrieve Termset and Term
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termSet = session.GetDefaultSiteCollectionTermStore().GetTermSet(_termSet1Id);
                var anchorTerm = termSet.GetTerm(_termId);
                clientContext.Load(termSet);
                clientContext.Load(anchorTerm);
                clientContext.ExecuteQueryRetry();

                // Retrieve List
                var list = clientContext.Web.Lists.GetById(_listId);
                clientContext.Load(list);
                clientContext.ExecuteQueryRetry();

                // Create field
                var fieldId = Guid.NewGuid();
                var fieldName = "Test_" + DateTime.Now.ToFileTime();
                TaxonomyFieldCreationInformation fieldCI = new TaxonomyFieldCreationInformation()
                {
                    Id = fieldId,
                    DisplayName = fieldName,
                    InternalName = fieldName,
                    Group = "Test Fields Group",
                    TaxonomyItem = anchorTerm
                };
                var field = list.CreateTaxonomyField(fieldCI);


                Assert.AreEqual(fieldId, field.Id, "Field IDs do not match.");
                Assert.AreEqual(fieldName, field.InternalName, "Field internal names do not match.");
                Assert.AreEqual("TaxonomyFieldType", field.TypeAsString, "Failed to create a TaxonomyField object.");
            }
        }
        #endregion

        #region Get taxonomy object tests
        [TestMethod()]
        public void GetTaxonomySessionTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var session = site.GetTaxonomySession();
                Assert.IsInstanceOfType(session, typeof(TaxonomySession), "Did not return TaxonomySession object");
            }
        }

        [TestMethod()]
        public void GetDefaultKeywordsTermStoreTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var termStore = site.GetDefaultKeywordsTermStore();
                Assert.IsInstanceOfType(termStore, typeof(TermStore), "Did not return TermStore object");
            }
        }

        [TestMethod()]
        public void GetDefaultSiteCollectionTermStoreTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var termStore = site.GetDefaultSiteCollectionTermStore();
                Assert.IsInstanceOfType(termStore, typeof(TermStore), "Did not return TermStore object");
            }
        }

        [TestMethod()]
        public void GetTermSetsByNameTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var termSetCollection = site.GetTermSetsByName(_termSetName1);
                Assert.IsInstanceOfType(termSetCollection, typeof(TermSetCollection), "Did not return TermSetCollection object");
                Assert.IsTrue(termSetCollection.AreItemsAvailable, "No terms available");
            }
        }

        [TestMethod()]
        public void GetTermGroupByNameTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var termGroup = site.GetTermGroupByName(_termGroupName);
                Assert.IsInstanceOfType(termGroup, typeof(TermGroup), "Did not return TermGroup object");
                Assert.AreEqual(_termGroupName, termGroup.Name, "Name does not match");
            }
        }

        [TestMethod()]
        public void GetTermGroupByIdTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var termGroup = site.GetTermGroupById(_termGroupId);
                Assert.IsInstanceOfType(termGroup, typeof(TermGroup), "Did not return TermGroup object");
                Assert.AreEqual(_termGroupId, termGroup.Id, "Name does not match");
            }
        }

        [TestMethod()]
        public void GetTermByNameTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var term = site.GetTermByName(_termSet1Id, _termName);
                Assert.IsInstanceOfType(term, typeof(Term), "Did not return Term object");
                Assert.AreEqual(_termName, term.Name, "Name does not match");
            }
        }

        [TestMethod()]
        public void GetTaxonomyItemByPathTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var path = _termGroupName + "|" + _termSetName1;
                var taxonomyItem = site.GetTaxonomyItemByPath(path);
                Assert.IsInstanceOfType(taxonomyItem, typeof(TaxonomyItem));
                Assert.AreEqual(_termSetName1, taxonomyItem.Name, "Did not return correct termset");

                path = _termGroupName + "|" + _termSetName1 + "|" + _termName;
                taxonomyItem = site.GetTaxonomyItemByPath(path);

                Assert.IsInstanceOfType(taxonomyItem, typeof(TaxonomyItem));
                Assert.AreEqual(_termName, taxonomyItem.Name, "Did not return correct term");
            }
        }
        #endregion

        #region Add term tests
        [TestMethod()]
        public void AddTermToTermsetTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var termName = "Test_Term_" + DateTime.Now.ToFileTime();
                var term = site.AddTermToTermset(_termSet1Id, termName);
                Assert.IsInstanceOfType(term, typeof(Term), "Did not return Term object");
                Assert.AreEqual(termName, term.Name, "Name does not match");
            }
        }

        [TestMethod()]
        public void AddTermToTermsetWithTermIdTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var termName = "Test_Term_" + DateTime.Now.ToFileTime();
                var termId = Guid.NewGuid();
                var term = site.AddTermToTermset(_termSet1Id, termName, termId);
                Assert.IsInstanceOfType(term, typeof(Term), "Did not return Term object");
                Assert.AreEqual(termName, term.Name, "Name does not match");
                Assert.AreEqual(termId, term.Id, "Id does not match");

            }
        }
        #endregion

        #region Import terms tests
        [TestMethod()]
        public void ImportTermsTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;

                var termName1 = "Test_Term_1" + DateTime.Now.ToFileTime();
                var termName2 = "Test_Term_2" + DateTime.Now.ToFileTime();

                List<string> termLines = new List<string>();
                termLines.Add(_termGroupName + "|" + _termSetName1 + "|" + termName1);
                termLines.Add(_termGroupName + "|" + _termSetName2 + "|" + termName2);
                site.ImportTerms(termLines.ToArray(), 1033, "|");

                var taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                var termGroup = termStore.Groups.GetByName(_termGroupName);
                var termSet1 = termGroup.TermSets.GetByName(_termSetName1);
                var termSet2 = termGroup.TermSets.GetByName(_termSetName2);
                var term1 = termSet1.Terms.GetByName(termName1);
                var term2 = termSet2.Terms.GetByName(termName2);
                clientContext.Load(term1);
                clientContext.Load(term2);
                clientContext.ExecuteQueryRetry();

                Assert.IsNotNull(term1);
                Assert.IsNotNull(term2);
            }
        }

        [TestMethod()]
        public void ImportTermsToTermStoreTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;

                var termName1 = "Test_Term_1" + DateTime.Now.ToFileTime();
                var termName2 = "Test_Term_2" + DateTime.Now.ToFileTime();

                List<string> termLines = new List<string>();
                termLines.Add(_termGroupName + "|" + _termSetName1 + "|" + termName1);
                termLines.Add(_termGroupName + "|" + _termSetName1 + "|" + termName2);

                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = session.GetDefaultSiteCollectionTermStore();
                site.ImportTerms(termLines.ToArray(), 1033, termStore, "|");

                var taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                var termGroup = termStore.Groups.GetByName(_termGroupName);
                var termSet = termGroup.TermSets.GetByName(_termSetName1);
                var term1 = termSet.Terms.GetByName(termName1);
                var term2 = termSet.Terms.GetByName(termName2);
                clientContext.Load(term1);
                clientContext.Load(term2);
                clientContext.ExecuteQueryRetry();

                Assert.IsNotNull(term1);
                Assert.IsNotNull(term2);
            }
        }

        [TestMethod()]
        public void HandleTermsWithCommaTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;

                var termName1 = "Comma,Comma";

                List<string> termLines = new List<string>();
                string termSrc1 = _termGroupName + "|" + _termSetName1 + "|\"" + termName1 + "\"";
                termLines.Add(termSrc1);

                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = session.GetDefaultSiteCollectionTermStore();
                site.ImportTerms(termLines.ToArray(), 1033, termStore, "|");

                var terms = site.ExportTermSet(_termSet1Id, false);
                string termDest1 = terms.SingleOrDefault(t => t.Contains(termName1));
                Assert.AreEqual(termSrc1, termDest1);
            }
        }

        [TestMethod()]
        public void HandleTermsWithQuotesTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;

                var termName1 = "\"Quotes and , comma\"";
                var termName2 = "Quote \" In the Middle";
                var termName3 = "\"Quote Start";
                var termName4 = "Quote End\"";
                var termName5 = "\"StartQuote \" MiddleQuote";

                List<string> termLines = new List<string>();
                string termSrc1 = _termGroupName + "|" + _termSetName1 + "|" + termName1;
                string termSrc2 = _termGroupName + "|" + _termSetName1 + "|" + termName2;
                string termSrc3 = _termGroupName + "|" + _termSetName1 + "|" + termName3;
                string termSrc4 = _termGroupName + "|" + _termSetName1 + "|" + termName4;
                string termSrc5 = _termGroupName + "|" + _termSetName1 + "|" + termName5;
                termLines.Add(termSrc1);
                termLines.Add(termSrc2);
                termLines.Add(termSrc3);
                termLines.Add(termSrc4);
                termLines.Add(termSrc5);

                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = session.GetDefaultSiteCollectionTermStore();
                site.ImportTerms(termLines.ToArray(), 1033, termStore, "|");

                var terms = site.ExportTermSet(_termSet1Id, false);
                string termDest1 = terms.SingleOrDefault(t => t.Contains(termName1));
                Assert.AreEqual(termSrc1, termDest1);
                string termDest2 = terms.SingleOrDefault(t => t.Contains(termName2));
                Assert.AreEqual(termSrc2, termDest2);
                string termDest3 = terms.SingleOrDefault(t => t.Contains(termName3));
                Assert.AreEqual(termSrc3, termDest3);
                string termDest4 = terms.SingleOrDefault(t => t.Contains(termName4));
                Assert.AreEqual(termSrc4, termDest4);
                string termDest5 = terms.SingleOrDefault(t => t.Contains(termName5));
                Assert.AreEqual(termSrc5, termDest5);
            }
        }

        [TestMethod()]
        public void ImportTermSetSampleShouldCreateSetTest()
        {
            var importSetId = Guid.NewGuid();
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                var termGroup = termStore.GetGroup(_termGroupId);

                // Act
                var termSet = termGroup.ImportTermSet(SampleTermSetPath, importSetId);
            }

            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                var createdSet = termStore.GetTermSet(importSetId);
                var allTerms = createdSet.GetAllTerms();
                var rootCollection = createdSet.Terms;
                clientContext.Load(createdSet);
                clientContext.Load(allTerms);
                clientContext.Load(rootCollection, ts => ts.Include(t => t.Name, t => t.Description, t => t.IsAvailableForTagging));
                clientContext.ExecuteQueryRetry();

                Assert.AreEqual("Political Geography", createdSet.Name);
                Assert.AreEqual("A sample term set, describing a simple political geography.", createdSet.Description);
                Assert.IsFalse(createdSet.IsOpenForTermCreation);
                Assert.AreEqual(12, allTerms.Count);

                Assert.AreEqual(1, rootCollection.Count);
                Assert.AreEqual("Continent", rootCollection[0].Name);
                Assert.AreEqual("One of the seven main land masses (Europe, Asia, Africa, North America, South America, Australia, and Antarctica)", rootCollection[0].Description);
                Assert.IsTrue(rootCollection[0].IsAvailableForTagging);
            }
        }

        [TestMethod()]
        public void ImportTermSetShouldUpdateSetTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                clientContext.Load(termStore, s => s.DefaultLanguage);
                clientContext.ExecuteQueryRetry();
                var lcid = termStore.DefaultLanguage;

                var termGroup = termStore.GetGroup(_termGroupId);
                var termSet = termGroup.CreateTermSet("Test Changes", UpdateTermSetId, lcid);
                termSet.Description = "Initial term set description";
                var retain1 = termSet.CreateTerm("Retain1", lcid, Guid.NewGuid());
                retain1.SetDescription("Test of deletes, adds and update", lcid);
                var update2 = retain1.CreateTerm("Update2", lcid, Guid.NewGuid());
                update2.SetDescription("Initial update2 description", lcid);
                var retain3 = update2.CreateTerm("Retain3", lcid, Guid.NewGuid());
                retain3.SetDescription("Test retaining same term", lcid);
                var delete2 = retain1.CreateTerm("Delete2", lcid, Guid.NewGuid());
                delete2.SetDescription("Term to delete", lcid);
                var delete3 = delete2.CreateTerm("Delete3", lcid, Guid.NewGuid());
                delete3.SetDescription("Child term to delete", lcid);
                clientContext.ExecuteQueryRetry();
            }

            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                var termGroup = termStore.GetGroup(_termGroupId);

                // Act
                var termSet = termGroup.ImportTermSet(SampleUpdateTermSetPath, UpdateTermSetId, synchroniseDeletions: true, termSetIsOpen: true);
            }

            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                var createdSet = termStore.GetTermSet(UpdateTermSetId);
                var allTerms = createdSet.GetAllTerms();
                var rootCollection = createdSet.Terms;
                clientContext.Load(createdSet);
                clientContext.Load(allTerms);
                clientContext.Load(rootCollection, ts => ts.Include(t => t.Name, t => t.Description, t => t.IsAvailableForTagging));
                clientContext.ExecuteQueryRetry();

                Assert.AreEqual("Updated term set description", createdSet.Description);
                Assert.IsTrue(createdSet.IsOpenForTermCreation);
                Assert.AreEqual(6, allTerms.Count);
                Assert.AreEqual(2, rootCollection.Count);

                var retain1Collection = rootCollection.First(t => t.Name == "Retain1").Terms;
                clientContext.Load(retain1Collection, ts => ts.Include(t => t.Name, t => t.Description, t => t.IsAvailableForTagging));
                clientContext.ExecuteQueryRetry();

                Assert.IsTrue(retain1Collection.Any(t => t.Name == "New2"));
                Assert.IsFalse(retain1Collection.Any(t => t.Name == "Delete2"));
                Assert.AreEqual("Changed description", retain1Collection.First(t => t.Name == "Update2").Description);
                Assert.IsFalse(retain1Collection.First(t => t.Name == "Update2").IsAvailableForTagging);
            }
        }

        [TestMethod()]
        public void ImportTermSetShouldUpdateByGuidTest()
        {
            var addedTermId = new Guid("{B564BD6F-21FF-4B60-9474-5E33F726DC6C}");
            var changedTermId = new Guid("{73DF85EE-313C-4485-A7B3-0FC3C17A7454}");

            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                clientContext.Load(termStore, s => s.DefaultLanguage);
                clientContext.ExecuteQueryRetry();
                var lcid = termStore.DefaultLanguage;

                var termGroup = termStore.GetGroup(_termGroupId);
                var termSet = termGroup.CreateTermSet("Test Guids", GuidTermSetId, lcid);
                termSet.Description = "Initial term set description";
                var retain1 = termSet.CreateTerm("Retain1", lcid, Guid.NewGuid());
                retain1.SetDescription("Retained term description", lcid);
                var toUpdate1 = termSet.CreateTerm("ToUpdate1", lcid, changedTermId);
                toUpdate1.SetDescription("Inital term description", lcid);
                clientContext.ExecuteQueryRetry();
            }

            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                var termGroup = termStore.GetGroup(_termGroupId);

                // Act
                var termSet = termGroup.ImportTermSet(SampleGuidTermSetPath, Guid.Empty);
            }

            using (var clientContext = TestCommon.CreateClientContext())
            {
                var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                var createdSet = termStore.GetTermSet(GuidTermSetId);
                var rootCollection = createdSet.Terms;
                clientContext.Load(createdSet);
                clientContext.Load(rootCollection, ts => ts.Include(t => t.Name, t => t.Id));
                clientContext.ExecuteQueryRetry();

                Assert.AreEqual("Updated Guids", createdSet.Name);
                Assert.AreEqual("Updated Test Guid term set description", createdSet.Description);
                Assert.AreEqual(3, rootCollection.Count);

                Assert.AreEqual(addedTermId, rootCollection.First(t => t.Name == "Added1").Id);
                Assert.IsTrue(rootCollection.Any(t => t.Name == "Retain1"));
                Assert.IsFalse(rootCollection.Any(t => t.Name == "ToUpdate1"));
                Assert.AreEqual("Changed1", rootCollection.First(t => t.Id == changedTermId).Name);
            }
        }
        #endregion

        #region Export term tests
        [TestMethod()]
        public void ExportTermSetTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var lines = site.ExportTermSet(_termSet1Id, false);
                Assert.IsTrue(lines.Any(), "No lines returned");
            }
        }

        [TestMethod()]
        public void ExportTermSetFromTermstoreTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                var termStore = session.GetDefaultSiteCollectionTermStore();

                var lines = site.ExportTermSet(_termSet1Id, false, termStore);
                Assert.IsTrue(lines.Any(), "No lines returned");
            }
        }

        [TestMethod()]
        public void ExportAllTermsTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var site = clientContext.Site;
                var lines = site.ExportAllTerms(false);
                Assert.IsTrue(lines.Any(), "No lines returned");
            }
        }
        #endregion
    }
}
