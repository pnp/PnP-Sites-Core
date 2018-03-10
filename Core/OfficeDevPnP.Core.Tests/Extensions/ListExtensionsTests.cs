using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Tests;
using OfficeDevPnP.Core.Entities;
namespace Microsoft.SharePoint.Client.Tests
{
    [TestClass()]
    public class ListExtensionsTests
    {
        private string _termGroupName; // For easy reference. Set in the Initialize method
        private string _termSetName; // For easy reference. Set in the Initialize method
        private string _termName; // For easy reference. Set in the Initialize method
        private string _textFieldName; // For easy reference. Set in the Initialize method
        private string _textFieldName2; // For easy reference. Set in the Initialize method

        private Guid _termGroupId;
        private Guid _termSetId;
        private Guid _termId;
        private Guid _textFieldId;
        private Guid _textFieldId2;

        private Guid _listId; // For easy reference
        private Guid webHookListId;

        #region Test initialize and cleanup
        [TestInitialize()]
        public void Initialize()
        {
            // Let's do webhook testing for both app-only as credential flows
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var list = clientContext.Web.CreateList(ListTemplateType.DocumentLibrary, "Test_list_" + DateTime.Now.ToFileTime(), false);
                list.EnsureProperty(p => p.Id);
                webHookListId = list.Id;
            }

            if (!TestCommon.AppOnlyTesting())
            {
                /*** Make sure that the user defined in the App.config has permissions to Manage Terms ***/
                // Create some taxonomy groups and terms
                using (var clientContext = TestCommon.CreateClientContext())
                {
                    _termGroupName = "Test_Group_" + DateTime.Now.ToFileTime();
                    _termSetName = "Test_Termset_" + DateTime.Now.ToFileTime();
                    _termName = "Test_Term_" + DateTime.Now.ToFileTime();
                    _textFieldName = "Test_Text_Field_" + DateTime.Now.ToFileTime();
                    _textFieldName2 = "Test_Text_Field2_" + DateTime.Now.ToFileTime();

                    _termGroupId = Guid.NewGuid();
                    _termSetId = Guid.NewGuid();
                    _termId = Guid.NewGuid();

                    // Termgroup
                    var taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                    var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                    var termGroup = termStore.CreateGroup(_termGroupName, _termGroupId);
                    clientContext.Load(termGroup);
                    clientContext.ExecuteQueryRetry();

                    // Termset
                    var termSet = termGroup.CreateTermSet(_termSetName, _termSetId, 1033);
                    clientContext.Load(termSet);
                    clientContext.ExecuteQueryRetry();

                    // Term
                    termSet.CreateTerm(_termName, 1033, _termId);
                    clientContext.ExecuteQueryRetry();

                    // List

                    _textFieldId = Guid.NewGuid();
                    _textFieldId2 = Guid.NewGuid();

                    var fieldCI = new FieldCreationInformation(FieldType.Text)
                    {
                        Id = _textFieldId,
                        InternalName = _textFieldName,
                        DisplayName = "Test Text Field",
                        Group = "Test Group"
                    };

                    var fieldCI2 = new FieldCreationInformation(FieldType.Text)
                    {
                        Id = _textFieldId2,
                        InternalName = _textFieldName2,
                        DisplayName = "Test Text Field 2",
                        Group = "Test Group"
                    };

                    var textfield = clientContext.Web.CreateField(fieldCI);
                    var textfield2 = clientContext.Web.CreateField(fieldCI2);

                    var list = clientContext.Web.CreateList(ListTemplateType.DocumentLibrary, "Test_list_" + DateTime.Now.ToFileTime(), false);

                    var field = clientContext.Web.Fields.GetByInternalNameOrTitle("TaxKeyword"); // Enterprise Metadata

                    list.Fields.Add(field);
                    list.Fields.Add(textfield);
                    list.Fields.Add(textfield2);

                    list.Update();
                    clientContext.Load(list);
                    clientContext.ExecuteQueryRetry();

                    _listId = list.Id;
                }
            }
            else
            {
                using (var clientContext = TestCommon.CreateClientContext())
                {
                    _textFieldName = "Test_Text_Field_" + DateTime.Now.ToFileTime();
                    _textFieldName2 = "Test_Text_Field2_" + DateTime.Now.ToFileTime();

                    // List
                    _textFieldId = Guid.NewGuid();
                    _textFieldId2 = Guid.NewGuid();

                    var fieldCI = new FieldCreationInformation(FieldType.Text)
                    {
                        Id = _textFieldId,
                        InternalName = _textFieldName,
                        DisplayName = "Test Text Field",
                        Group = "Test Group"
                    };

                    var fieldCI2 = new FieldCreationInformation(FieldType.Text)
                    {
                        Id = _textFieldId2,
                        InternalName = _textFieldName2,
                        DisplayName = "Test Text Field 2",
                        Group = "Test Group"
                    };

                    var textfield = clientContext.Web.CreateField(fieldCI);
                    var textfield2 = clientContext.Web.CreateField(fieldCI2);

                    var list = clientContext.Web.CreateList(ListTemplateType.DocumentLibrary, "Test_list_" + DateTime.Now.ToFileTime(), false);

                    list.Fields.Add(textfield);
                    list.Fields.Add(textfield2);

                    list.Update();
                    clientContext.Load(list);
                    clientContext.ExecuteQueryRetry();

                    _listId = list.Id;
                }

            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            // Let's do webhook testing for both app-only as credential flows
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var list = clientContext.Web.Lists.GetById(webHookListId);
                list.DeleteObject();
                clientContext.ExecuteQueryRetry();
            }

            using (var clientContext = TestCommon.CreateClientContext())
            {
                if (!TestCommon.AppOnlyTesting())
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
                    }
                    termGroup.DeleteObject(); // Will delete underlying termset
                    clientContext.ExecuteQueryRetry();
                }

                // Clean up list
                var list = clientContext.Web.Lists.GetById(_listId);
                list.DeleteObject();
                clientContext.ExecuteQueryRetry();

                // Clean up fields
                var fields = clientContext.LoadQuery(clientContext.Web.Fields);
                clientContext.ExecuteQueryRetry();
                var testFields = fields.Where(f => f.InternalName.StartsWith("Test_", StringComparison.OrdinalIgnoreCase));
                foreach (var field in testFields)
                {
                    field.DeleteObject();
                }
                clientContext.ExecuteQueryRetry();
            }
        }
        #endregion

        #region Create list tests
        [TestMethod()]
        public void CreateListTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var listName = "Test_list_" + DateTime.Now.ToFileTime();

                //Create List
                var web = clientContext.Web;
                web.CreateList(ListTemplateType.GenericList, listName, false);

                //Get List
                var list = web.GetListByTitle(listName);

                Assert.IsNotNull(list);
                Assert.AreEqual(listName, list.Title);

                //Delete List
                list.DeleteObject();
                clientContext.ExecuteQueryRetry();
            }
        }
        #endregion

        #region List Existence tests
        [TestMethod]
        public void ListExistsByGuidTest()
        {
            var listName = "samplelist_" + DateTime.Now.ToFileTime();
            var listGuid = Guid.NewGuid();
            using (var clientContext = TestCommon.CreateClientContext())
            {

                var list = clientContext.Web.CreateList(
                    ListTemplateType.GenericList,
                    listName,
                    false);


                clientContext.Load<List>(list, l => l.Id);
                clientContext.ExecuteQueryRetry();

                Assert.IsNotNull(list);
                Assert.IsTrue(clientContext.Web.ListExists(list.Id));

                //Delete List
                list.DeleteObject();
                clientContext.ExecuteQueryRetry();

            }
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ListExistsWithEmtpyTitleParameterTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                clientContext.Web.ListExists(string.Empty);
            }
        }

        [TestMethod]
        public void ListExistsByTitleTest()
        {
            var listName = "samplelist_" + DateTime.Now.ToFileTime();
            var listGuid = Guid.NewGuid();
            using (var clientContext = TestCommon.CreateClientContext())
            {

                var list = clientContext.Web.CreateList(
                    ListTemplateType.GenericList,
                    listName,
                    false);

                Assert.IsNotNull(list);
                Assert.IsTrue(clientContext.Web.ListExists(listName));

                //Delete List
                list.DeleteObject();
                clientContext.ExecuteQueryRetry();

            }
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ListExistsByUrlPathIsNullParamTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                clientContext.Web.ListExists((Uri)null);
            }
        }

        [TestMethod]
        public void ListExistsByUrlPathParamTest()
        {
            var listName = "samplelist_" + DateTime.Now.ToFileTime();
            var siteRelativePath = $"Lists/{listName}";

            using (var clientContext = TestCommon.CreateClientContext())
            {
                var list = clientContext.Web.CreateList(
                    ListTemplateType.GenericList,
                    listName,
                    false);

                Assert.IsNotNull(list);
                Assert.IsTrue(clientContext.Web.ListExists(new Uri(siteRelativePath, UriKind.Relative)));

                //Delete List
                list.DeleteObject();
                clientContext.ExecuteQueryRetry();
            }
        }
        #endregion  

        #region Get Lists/Library tests

        public void GetPagesLibraryTest()
        {
            const string publishingWebFeatureId = "22a9ef51-737b-4ff2-9346-694633fe4416";
            using (var clientContext = TestCommon.CreateClientContext())
            {
                if (!clientContext.Web.IsFeatureActive(new Guid(publishingWebFeatureId)))
                {
                    Assert.Inconclusive("Can't execute GetPagesLibraryTest on a web without activated Publishing feature.");
                }

                var web = clientContext.Web;
                var pages = web.GetPagesLibrary();

                Assert.IsNotNull(pages);
            }
        }

        #endregion

        #region Default value tests
        [TestMethod()]
        public void SetDefaultColumnValuesTest()
        {
            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Taxonomy tests are not supported when testing using app-only");
            }

            using (var clientContext = TestCommon.CreateClientContext())
            {
                TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                List<IDefaultColumnValue> defaultValues = new List<IDefaultColumnValue>();

                var defaultColumnValue = new DefaultColumnTermValue();

                defaultColumnValue.FieldInternalName = "TaxKeyword"; // Enterprise metadata field, should be present on the list
                defaultColumnValue.FolderRelativePath = "/"; // Root Folder
                var term = taxSession.GetTerm(_termId);
                defaultColumnValue.Terms.Add(term);
                defaultValues.Add(defaultColumnValue);

                var testDefaultValue = new DefaultColumnTextValue();
                testDefaultValue.Text = "Bla";
                testDefaultValue.FieldInternalName = _textFieldName;
                testDefaultValue.FolderRelativePath = "/"; // Root folder

                defaultValues.Add(testDefaultValue);

                var list = clientContext.Web.Lists.GetById(_listId);

                list.SetDefaultColumnValues(defaultValues);
            }
        }

        [TestMethod()]
        public void ClearOneDefaultColumnValuesTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // arrange
                List<IDefaultColumnValue> defaultValues = new List<IDefaultColumnValue>();

                var testFooDefaultValue = new DefaultColumnTextValue();
                testFooDefaultValue.Text = "Foo";
                testFooDefaultValue.FieldInternalName = _textFieldName;
                testFooDefaultValue.FolderRelativePath = "/"; // Root folder

                var testBarDefaultValue = new DefaultColumnTextValue();
                testBarDefaultValue.Text = "Bar";
                testBarDefaultValue.FieldInternalName = _textFieldName2;
                testBarDefaultValue.FolderRelativePath = "/"; // Root folder

                defaultValues.Add(testFooDefaultValue);
                defaultValues.Add(testBarDefaultValue);

                var list = clientContext.Web.Lists.GetById(_listId);

                list.SetDefaultColumnValues(defaultValues);
                var result = list.GetDefaultColumnValues();
                var itemToRemove = result.First(d => d["Field"] == _textFieldName);
                result.Remove(itemToRemove);
                var expected = result;

                // act
                list.ClearDefaultColumnValues(new List<IDefaultColumnValue> { testFooDefaultValue });
                var actual = list.GetDefaultColumnValues();

                // assert
                CollectionAssert.AreEqual(actual[0], expected[0]);
            }
        }

        [TestMethod()]
        public void ClearAllDefaultColumnValuesTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // arrange
                List<IDefaultColumnValue> defaultValues = new List<IDefaultColumnValue>();

                var testFooDefaultValue = new DefaultColumnTextValue();
                testFooDefaultValue.Text = "Foo";
                testFooDefaultValue.FieldInternalName = _textFieldName;
                testFooDefaultValue.FolderRelativePath = "/"; // Root folder

                var testBarDefaultValue = new DefaultColumnTextValue();
                testBarDefaultValue.Text = "Bar";
                testBarDefaultValue.FieldInternalName = _textFieldName2;
                testBarDefaultValue.FolderRelativePath = "/"; // Root folder

                defaultValues.Add(testFooDefaultValue);
                defaultValues.Add(testBarDefaultValue);

                var list = clientContext.Web.Lists.GetById(_listId);

                list.SetDefaultColumnValues(defaultValues);

                // act
                list.ClearDefaultColumnValues();
                var actual = list.GetDefaultColumnValues();

                // assert
                Assert.IsNull(actual);
            }
        }

        [TestMethod()]
        public void OverwriteDefaultColumnValuesTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                // arrange
                List<IDefaultColumnValue> defaultValues = new List<IDefaultColumnValue>();

                var testFooDefaultValue = new DefaultColumnTextValue();
                testFooDefaultValue.Text = "Foo";
                testFooDefaultValue.FieldInternalName = _textFieldName;
                testFooDefaultValue.FolderRelativePath = "/"; // Root folder

                var testBarDefaultValue = new DefaultColumnTextValue();
                testBarDefaultValue.Text = "Bar";
                testBarDefaultValue.FieldInternalName = _textFieldName2;
                testBarDefaultValue.FolderRelativePath = "/"; // Root folder

                defaultValues.Add(testFooDefaultValue);
                defaultValues.Add(testBarDefaultValue);

                var list = clientContext.Web.Lists.GetById(_listId);

                list.SetDefaultColumnValues(defaultValues);
                var result = list.GetDefaultColumnValues();
                var itemToRemove = result.First(d => d["Field"] == _textFieldName);
                result.Remove(itemToRemove);
                var expected = result;

                // act
                list.SetDefaultColumnValues(new List<IDefaultColumnValue> { testBarDefaultValue }, true);
                var actual = list.GetDefaultColumnValues();

                // assert
                CollectionAssert.AreEqual(actual[0], expected[0]);
            }
        }
        #endregion

        #region Webhooks tests
#if !ONPREMISES
        [TestMethod]
        public void AddWebhookTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                clientContext.ExecuteQueryRetry();

                WebhookSubscription expectedSubscription = new WebhookSubscription()
                {
                    ExpirationDateTime = DateTime.Today.AddMonths(3),
                    NotificationUrl = TestCommon.TestWebhookUrl,
                    Resource = TestCommon.DevSiteUrl + string.Format("/_api/lists('{0}')", webHookListId)
                };
                WebhookSubscription actualSubscription = testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, 3);

                // Compare properties of expected and actual
                Assert.IsTrue(Equals(expectedSubscription.ClientState, actualSubscription.ClientState) 
                    && Equals(expectedSubscription.ExpirationDateTime.Date, actualSubscription.ExpirationDateTime.Date)
                    && Equals(expectedSubscription.NotificationUrl, actualSubscription.NotificationUrl)
                    && expectedSubscription.Resource.Contains(actualSubscription.Resource));
            }
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void AddWebhookWithInvalidExpirationDateTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                clientContext.ExecuteQueryRetry();
                
                testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, 12);
            }
        }

        [TestMethod]
        public void AddWebhookWithVeryLastValidExpirationDateTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                clientContext.ExecuteQueryRetry();
                
                DateTime veryLastValidExpiration = DateTime.UtcNow.AddDays(180);

                WebhookSubscription expectedSubscription = new WebhookSubscription()
                {
                    ExpirationDateTime = veryLastValidExpiration,
                    NotificationUrl = TestCommon.TestWebhookUrl,
                    Resource = TestCommon.DevSiteUrl + string.Format("/_api/lists('{0}')", webHookListId)
                };
                WebhookSubscription actualSubscription = testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, veryLastValidExpiration);

                // Compare properties of expected and actual
                Assert.IsTrue(Equals(expectedSubscription.ClientState, actualSubscription.ClientState)
                    && Equals(expectedSubscription.ExpirationDateTime.Date, actualSubscription.ExpirationDateTime.Date)
                    && Equals(expectedSubscription.NotificationUrl, actualSubscription.NotificationUrl)
                    && expectedSubscription.Resource.Contains(actualSubscription.Resource));
            }
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void AddWebhookWithBarelyInvalidExpirationDateTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                clientContext.ExecuteQueryRetry();

                DateTime barelyInvalidExpiration = DateTime.UtcNow.AddDays(180).AddMinutes(1);

                testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, barelyInvalidExpiration);
            }
        }

        [TestMethod]
        public void UpdateWebhookTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                clientContext.ExecuteQueryRetry();

                WebhookSubscription actualSubscription = testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, 3);

                // Change the expiration time
                actualSubscription.ExpirationDateTime = actualSubscription.ExpirationDateTime.AddDays(7);

                bool result = testList.UpdateWebhookSubscription(actualSubscription);

                Assert.IsTrue(result);
            }
        }

        [TestMethod]
        public void RemoveWebhookTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                clientContext.ExecuteQueryRetry();

                WebhookSubscription actualSubscription = testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, 3);

                bool result = testList.RemoveWebhookSubscription(actualSubscription.Id);

                Assert.IsTrue(result);
            }
        }

        [TestMethod]
        public void GetAllWebhookSubscriptionsTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                clientContext.ExecuteQueryRetry();

                WebhookSubscription createdSubscription = testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, 3);

                IList<WebhookSubscription> actualSubscriptions = Task.Run(()=> testList.GetWebhookSubscriptionsAsync()).Result;

                Assert.IsTrue(actualSubscriptions.Count > 0);
            }
        }

        [TestMethod]
        public async Task GetAllWebhookSubscriptionsTestAsync()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var testList = clientContext.Web.Lists.GetById(webHookListId);
                clientContext.Load(testList);
                await clientContext.ExecuteQueryRetryAsync();

                WebhookSubscription createdSubscription = testList.AddWebhookSubscription(TestCommon.TestWebhookUrl, 3);

                IList<WebhookSubscription> actualSubscriptions = await testList.GetWebhookSubscriptionsAsync();

                Assert.IsTrue(actualSubscriptions.Count > 0);
            }
        }

#endif
        #endregion

    }
}
