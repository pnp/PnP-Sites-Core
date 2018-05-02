using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Xml.Linq;
using Microsoft.SharePoint.Client.Taxonomy;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectListInstanceTests
    {
        private const string ElementSchema = @"<Field xmlns=""http://schemas.microsoft.com/sharepoint/v3"" Name=""DemoField"" StaticName=""DemoField"" DisplayName=""Test Field"" Type=""Text"" ID=""{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}"" Group=""PnP"" Required=""true""/>";
        private Guid fieldId = Guid.Parse("{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}");
        private Guid termGroupId = Guid.Empty;

        private const string CalculatedFieldElementSchema = @"<Field Name=""CalculatedField"" StaticName=""CalculatedField"" DisplayName=""Test Calculated Field"" Type=""Calculated"" ResultType=""Text"" ID=""{D1A33456-9FEB-4D8E-AFFA-177EACCE4B70}"" Group=""PnP"" ReadOnly=""TRUE"" ><Formula>=DemoField&amp;""DemoField""</Formula><FieldRefs><FieldRef Name=""DemoField"" ID=""{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}"" /></FieldRefs></Field>";
        private const string TokenizedCalculatedFieldElementSchema = @"<Field Name=""CalculatedField"" StaticName=""CalculatedField"" DisplayName=""Test Calculated Field"" Type=""Calculated"" ResultType=""Text"" ID=""{D1A33456-9FEB-4D8E-AFFA-177EACCE4B70}"" Group=""PnP"" ReadOnly=""TRUE"" ><Formula>=[{fieldtitle:DemoField}]&amp;""DemoField""</Formula></Field>";
        private Guid calculatedFieldId = Guid.Parse("{D1A33456-9FEB-4D8E-AFFA-177EACCE4B70}");


        private string listName;
        private string datarowListName;
        [TestInitialize]
        public void Initialize()
        {
            listName = string.Format("Test_{0}", DateTime.Now.Ticks);
            datarowListName = $"DataRowTest_{DateTime.Now.Ticks}";

        }
        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                bool isDirty = false;

                var list = ctx.Web.GetListByUrl(string.Format("lists/{0}", listName));
                if (list == null)
                    list = ctx.Web.GetListByUrl(listName);
                if (list != null)
                {
                    list.DeleteObject();
                    isDirty = true;
                }

                // Clean all data row test list instances, also after a previous test case failed.
                DeleteDataRowLists(ctx);

                // first delete content types
                var contentTypes = ctx.LoadQuery(ctx.Web.ContentTypes);
                ctx.ExecuteQueryRetry();
                var testContentTypes = contentTypes.Where(l => l.Name.StartsWith("Test_", StringComparison.OrdinalIgnoreCase));
                foreach (var ctype in testContentTypes)
                {
                    ctype.DeleteObject();
                    isDirty = true;
                }

                var field = ctx.Web.GetFieldById<FieldText>(fieldId); // Guid matches ID in field caml.
                var calculatedField = ctx.Web.GetFieldById<FieldCalculated>(calculatedFieldId); // Guid matches ID in field caml.

                if (field != null)
                {
                    field.DeleteObject();
                    isDirty = true;
                }
                if (calculatedField != null)
                {
                    calculatedField.DeleteObject();
                    isDirty = true;
                }

                if (isDirty)
                {
                    ctx.ExecuteQueryRetry();
                }

                if (!TestCommon.AppOnlyTesting())
                {
                    // Clean up Taxonomy
                    if (!Guid.Empty.Equals(termGroupId))
                    {
                        var taxSession = TaxonomySession.GetTaxonomySession(ctx);
                        var termStore = taxSession.GetDefaultSiteCollectionTermStore();
                        var termGroup = termStore.GetGroup(termGroupId);
                        ctx.ExecuteQueryRetry();
                        isDirty = false;
                        if (!termGroup.ServerObjectIsNull.Value)
                        {
                            var termSets = termGroup.TermSets;
                            ctx.Load(termSets);
                            ctx.ExecuteQueryRetry();
                            foreach (var termSet in termSets)
                            {
                                termSet.DeleteObject();
                            }
                            termGroup.DeleteObject();
                            isDirty = true;
                        }
                        if (isDirty)
                        {
                            ctx.ExecuteQueryRetry();
                        }
                    }
                }
            }
        }

        private void DeleteDataRowLists(ClientContext cc)
        {
            cc.Load(cc.Web.Lists, f => f.Include(t => t.Title));
            cc.ExecuteQueryRetry();

            foreach (var list in cc.Web.Lists.ToList())
            {
                if (list.Title.StartsWith("DataRowTest_"))
                {
                    list.DeleteObject();
                }
            }
            cc.ExecuteQueryRetry();
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Taxonomy tests are not supported when testing using app-only");
            }

            var template = new ProvisioningTemplate();
            var listInstance = new Core.Framework.Provisioning.Model.ListInstance();

            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;
            listInstance.FieldRefs.Add(new FieldRef() { Id = new Guid("23f27201-bee3-471e-b2e7-b64fd8b7ca38") });

            using (var ctx = TestCommon.CreateClientContext())
            {
                //Create term
                var taxSession = TaxonomySession.GetTaxonomySession(ctx);
                var termStore = taxSession.GetDefaultSiteCollectionTermStore();

                // Termgroup
                termGroupId = Guid.NewGuid();
                var termGroup = termStore.CreateGroup("Test_Group_" + DateTime.Now.ToFileTime(), termGroupId);
                ctx.Load(termGroup);

                var termSet = termGroup.CreateTermSet("Test_Termset_" + DateTime.Now.ToFileTime(), Guid.NewGuid(), 1033);
                ctx.Load(termSet);

                Guid termId = Guid.NewGuid();
                string termName = "Test_Term_" + DateTime.Now.ToFileTime();

                termSet.CreateTerm(termName, 1033, termId);

                Dictionary<string, string> dataValues = new Dictionary<string, string>();
                dataValues.Add("Title", "Test");
                dataValues.Add("TaxKeyword", $"{termName}|{termId.ToString()}");
                DataRow dataRow = new DataRow(dataValues);

                listInstance.DataRows.Add(dataRow);

                template.Lists.Add(listInstance);


                var parser = new TokenParser(ctx.Web, template);

                // Create the List
                parser = new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                parser = new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                // Load DataRows
                new ObjectListInstanceDataRows().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var items = list.GetItems(CamlQuery.CreateAllItemsQuery());
                ctx.Load(items, itms => itms.Include(item => item["Title"], i => i["TaxKeyword"]));
                ctx.ExecuteQueryRetry();

                Assert.IsTrue(items.Count == 1);
                Assert.IsTrue(items[0]["Title"].ToString() == "Test");

                //Validate taxonomy field data
                var value = items[0]["TaxKeyword"] as TaxonomyFieldValueCollection;
                Assert.IsNotNull(value);
                Assert.IsTrue(value[0].WssId > 0, "Term WSS ID not set correctly");
                Assert.AreEqual(termName, value[0].Label, "Term label not set correctly");
                Assert.AreEqual(termId.ToString(), value[0].TermGuid, "Term GUID not set correctly");

            }
        }

        [TestMethod]
        public void CanCreateEntities()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template = new ProvisioningTemplate();
                template = new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.Export).ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.Lists.Any());
            }
        }

        [TestMethod]
        public void CanTokensBeUsedInListInstance()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                // Create list instance
                var template = new ProvisioningTemplate();

                var listUrl = string.Format("lists/{0}", listName);
                var listTitle = listName + "_Title";
                var listDesc = listName + "_Description";
                template.Parameters.Add("listTitle", listTitle);
                template.Parameters.Add("listDesc", listDesc);

                template.Lists.Add(new Core.Framework.Provisioning.Model.ListInstance
                {
                    Url = listUrl,
                    Title = "{parameter:listTitle}",
                    Description = "{parameter:listDesc}",
                    TemplateType = (int)ListTemplateType.GenericList
                });

                ctx.Web.ApplyProvisioningTemplate(template);

                var list = ctx.Web.GetListByUrl(listUrl, l => l.Title, l => l.Description);
                Assert.IsNotNull(list);
                Assert.AreEqual(listTitle, list.Title);
                Assert.AreEqual(listDesc, list.Description);

                // Update list instance
                var updatedTemplate = new ProvisioningTemplate();

                var updatedTitle = listName + "_UpdatedTitle";
                var updatedDesc = listName + "_UpdatedDescription";
                updatedTemplate.Parameters.Add("listTitle", updatedTitle);
                updatedTemplate.Parameters.Add("listDesc", updatedDesc);

                updatedTemplate.Lists.Add(new Core.Framework.Provisioning.Model.ListInstance
                {
                    Url = listUrl,
                    Title = "{parameter:listTitle}",
                    Description = "{parameter:listDesc}",
                    TemplateType = (int)ListTemplateType.GenericList
                });

                ctx.Web.ApplyProvisioningTemplate(updatedTemplate);

                var updatedList = ctx.Web.GetListByUrl(listUrl, l => l.Title, l => l.Description);
                Assert.AreEqual(updatedTitle, updatedList.Title);
                Assert.AreEqual(updatedDesc, updatedList.Description);
            }
        }

        [TestMethod]
        public void FolderContentTypeShouldNotBeRemovedFromProvisionedDocumentLibraries()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var listInstance = new Core.Framework.Provisioning.Model.ListInstance();
                listInstance.Url = listName;
                listInstance.Title = listName;
                listInstance.TemplateType = (int)ListTemplateType.DocumentLibrary;
                listInstance.ContentTypesEnabled = true;
                listInstance.RemoveExistingContentTypes = true;
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding { ContentTypeId = BuiltInContentTypeId.DublinCoreName, Default = true });
                var template = new ProvisioningTemplate();
                template.Lists.Add(listInstance);

                ctx.Web.ApplyProvisioningTemplate(template);

                var list = ctx.Web.GetListByUrl(listName);
                var contentTypes = list.EnsureProperty(l => l.ContentTypes);
                Assert.IsTrue(contentTypes.Any(ct => ct.StringId.StartsWith(BuiltInContentTypeId.Folder + "00")), "Folder content type should not be removed from a document library.");
            }

        }

        [TestMethod]
        public void DefaultContentTypeShouldBeRemovedFromProvisionedAssetLibraries()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                // Arrange
                var listInstance = new Core.Framework.Provisioning.Model.ListInstance();
                listInstance.Url = $"lists/{listName}";
                listInstance.Title = listName;
                // An asset must be created by using the 
                // template type AND the template feature id
                listInstance.TemplateType = 851;
                listInstance.TemplateFeatureID = new Guid("4bcccd62-dcaf-46dc-a7d4-e38277ef33f4");
                // Also attachements are not allowed on an asset list
                listInstance.EnableAttachments = false;
                listInstance.ContentTypesEnabled = true;
                listInstance.RemoveExistingContentTypes = true;
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding
                {
                    ContentTypeId = BuiltInContentTypeId.DublinCoreName,
                    Default = true
                });
                var template = new ProvisioningTemplate();
                template.Lists.Add(listInstance);

                // Act
                ctx.Web.ApplyProvisioningTemplate(template);
                var list = ctx.Web.GetListByUrl(listInstance.Url);
                var contentTypes = list.EnsureProperty(l => l.ContentTypes);
                // Assert
                // Asset list should only have the custom content type we defined
                // and the folder content type
                Assert.AreEqual(contentTypes.Count, 2);
            }

        }

#if !NETSTANDARD2_0
        [TestMethod]
        public void UpdatedListTitleShouldBeAvailableAsToken()
        {
            var listUrl = string.Format("lists/{0}", listName);
            var listId = "";

            // Create the initial list
            using (var ctx = TestCommon.CreateClientContext())
            {
                var list = ctx.Web.Lists.Add(new ListCreationInformation() { Title = listName, TemplateType = (int)ListTemplateType.GenericList, Url = listUrl });
                list.EnsureProperty(l => l.Id);
                ctx.ExecuteQueryRetry();
                listId = list.Id.ToString();
            }

            // Update list Title using a provisioning template 
            // - Using a clean clientcontext to catch all possible "property not loaded" problems
            using (var ctx = TestCommon.CreateClientContext())
            {
                var updatedListTitle = listName + "_edit";
                var template = new ProvisioningTemplate();
                var listInstance = new Core.Framework.Provisioning.Model.ListInstance();
                listInstance.Url = listUrl;
                listInstance.Title = updatedListTitle;
                listInstance.TemplateType = (int)ListTemplateType.GenericList;
                template.Lists.Add(listInstance);
                var mockProviderType = typeof(MockProviderForListInstanceTests);
                var providerConfig = "{listid:" + updatedListTitle + "}+{listurl:" + updatedListTitle + "}";
                template.Providers.Add(new Provider() { Assembly = mockProviderType.Assembly.FullName, Type = mockProviderType.FullName, Enabled = true, Configuration = providerConfig });
                ctx.Web.ApplyProvisioningTemplate(template);
            }

            // Verify that tokens have been replaced
            var expectedConfig = string.Format("{0}+{1}", listId, listUrl).ToLower();
            Assert.AreEqual(expectedConfig, MockProviderForListInstanceTests.ConfigurationData.ToLower(), "Updated list title is not available as a token.");
        }

        class MockProviderForListInstanceTests : OfficeDevPnP.Core.Framework.Provisioning.Extensibility.IProvisioningExtensibilityProvider
        {
            public static string ConfigurationData { get; private set; }
            public void ProcessRequest(ClientContext ctx, ProvisioningTemplate template, string configurationData)
            {
                ConfigurationData = configurationData;
            }
        }
#endif
        [TestMethod]
        public void CanProvisionCalculatedFieldRefInListInstance()
        {
            var template = new ProvisioningTemplate();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = TokenizedCalculatedFieldElementSchema });

            var listInstance = new ListInstance();
            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;

            var referencedField = new FieldRef();
            referencedField.Id = fieldId;
            listInstance.FieldRefs.Add(referencedField);

            var calculatedFieldRef = new FieldRef();
            calculatedFieldRef.Id = calculatedFieldId;
            listInstance.FieldRefs.Add(calculatedFieldRef);
            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f);
                Assert.IsInstanceOfType(f, typeof(FieldCalculated));
                Assert.IsFalse(f.Formula.Contains('#') || f.Formula.Contains('?'), "Calculated field was not provisioned properly");
            }
        }

        [TestMethod]
        public void CanUpdateCalculatedFieldRefInListInstance()
        {
            var template = new ProvisioningTemplate();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = TokenizedCalculatedFieldElementSchema });

            var listInstance = new ListInstance();
            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;

            var referencedField = new FieldRef();
            referencedField.Id = fieldId;
            listInstance.FieldRefs.Add(referencedField);

            var calculatedFieldRef = new FieldRef();
            calculatedFieldRef.Id = calculatedFieldId;
            listInstance.FieldRefs.Add(calculatedFieldRef);
            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f1 = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f1);
                Assert.IsInstanceOfType(f1, typeof(FieldCalculated));
                Assert.IsFalse(f1.Formula.Contains('#') || f1.Formula.Contains('?'), "Calculated field was not provisioned properly the first time");

                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var f2 = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(f2);
                Assert.IsInstanceOfType(f2, typeof(FieldCalculated));
                Assert.IsFalse(f2.Formula.Contains('#') || f2.Formula.Contains('?'), "Calculated field was not provisioned properly the second time");
            }
        }

        [TestMethod]
        public void CanProvisionCalculatedFieldInListInstance()
        {
            var template = new ProvisioningTemplate();
            var listInstance = new ListInstance();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });

            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;

            var referencedField = new FieldRef();
            referencedField.Id = fieldId;
            listInstance.FieldRefs.Add(referencedField);

            var calculatedField = new Core.Framework.Provisioning.Model.Field();
            calculatedField.SchemaXml = TokenizedCalculatedFieldElementSchema;
            listInstance.Fields.Add(calculatedField);

            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f);
                Assert.IsInstanceOfType(f, typeof(FieldCalculated));
                Assert.IsFalse(f.Formula.Contains('#') || f.Formula.Contains('?'), "Calculated field was not provisioned properly");
            }
        }

        [TestMethod]
        public void CanProvisionCalculatedFieldLocallyInListInstance()
        {
            //This test will fail as tokens does not support this scenario.
            //The test serves as a reminder that this is not supported and needs to be fixed in a future release.
            var template = new ProvisioningTemplate();
            var listInstance = new ListInstance();

            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;
            var referencedField = new Core.Framework.Provisioning.Model.Field();
            referencedField.SchemaXml = ElementSchema;
            listInstance.Fields.Add(referencedField);
            var calculatedField = new Core.Framework.Provisioning.Model.Field();
            calculatedField.SchemaXml = TokenizedCalculatedFieldElementSchema;
            listInstance.Fields.Add(calculatedField);
            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f);
                Assert.IsInstanceOfType(f, typeof(FieldCalculated));
                Assert.IsFalse(f.Formula.Contains('#') || f.Formula.Contains('?'), "Calculated field was not provisioned properly");
            }
        }

        [TestMethod]
        public void CanUpdateCalculatedFieldInListInstance()
        {
            var template = new ProvisioningTemplate();
            var listInstance = new ListInstance();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });

            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;

            var referencedField = new FieldRef();
            referencedField.Id = fieldId;
            listInstance.FieldRefs.Add(referencedField);

            var calculatedField = new Core.Framework.Provisioning.Model.Field();
            calculatedField.SchemaXml = TokenizedCalculatedFieldElementSchema;
            listInstance.Fields.Add(calculatedField);

            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f1 = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f1);
                Assert.IsInstanceOfType(f1, typeof(FieldCalculated));
                Assert.IsFalse(f1.Formula.Contains('#') || f1.Formula.Contains('?'), "Calculated field was not provisioned properly the first time");

                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var f2 = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(f2);
                Assert.IsInstanceOfType(f2, typeof(FieldCalculated));
                Assert.IsFalse(f2.Formula.Contains('#') || f2.Formula.Contains('?'), "Calculated field was not provisioned properly the second time");
            }
        }

        [TestMethod]
        public void CanExtractCalculatedFieldFromListInstance()
        {
            var template = new ProvisioningTemplate();
            var listInstance = new ListInstance();

            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });

            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;

            var referencedField = new FieldRef();
            referencedField.Id = fieldId;
            listInstance.FieldRefs.Add(referencedField);

            var calculatedField = new Core.Framework.Provisioning.Model.Field();
            calculatedField.SchemaXml = TokenizedCalculatedFieldElementSchema;
            listInstance.Fields.Add(calculatedField);
            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f);
                Assert.IsInstanceOfType(f, typeof(FieldCalculated));
                Assert.IsFalse(f.Formula.Contains('#') || f.Formula.Contains('?'), "Calculated field was not provisioned properly");

                var extractedTemplate = new ProvisioningTemplate();
                var provisioningTemplateCreationInformation = new ProvisioningTemplateCreationInformation(ctx.Web);
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ExtractObjects(ctx.Web, extractedTemplate, provisioningTemplateCreationInformation);
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ExtractObjects(ctx.Web, extractedTemplate, provisioningTemplateCreationInformation);

                XElement fieldElement = XElement.Parse(extractedTemplate.Lists.First(l => l.Title == listName).Fields.First(cf => Guid.Parse(XElement.Parse(cf.SchemaXml).Attribute("ID").Value).Equals(calculatedFieldId)).SchemaXml);
                var formula = fieldElement.Descendants("Formula").FirstOrDefault();

                Assert.AreEqual(@"=[{fieldtitle:DemoField}]&""DemoField""", formula.Value, true, "Calculated field formula is not extracted properly");
            }
        }

        [TestMethod]
        public void DataRowsAreBeingSkippedIfAlreadyInplace()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var template = new ProvisioningTemplate();
                template.TemplateCultureInfo = "1033";
                var listinstance = new ListInstance()
                {
                    Title = datarowListName,
                    Url = $"lists/{datarowListName}",
                    TemplateType = 100,
                };
                listinstance.Fields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = $@"<Field Type=""Text"" DisplayName=""Key"" Required=""FALSE"" EnforceUniqueValues=""FALSE"" Indexed=""FALSE"" MaxLength=""255"" ID=""{(Guid.NewGuid().ToString("B"))}"" StaticName=""Key"" Name=""Key"" />" });

                var datarows = new List<DataRow>()
                {
                    new DataRow(new Dictionary<string, string>{ { "Title", "Test -1-"}, { "Key", "1" } }, "1" ),
                    new DataRow(new Dictionary<string,string>{{ "Title" ,"Test -2-"}, { "Key", "2" } }, "2"),
                    new DataRow(new Dictionary<string,string>{{ "Title" ,"Test -3-"}, { "Key", "3" } }, "3")
                };
                listinstance.DataRows.AddRange(datarows);
                template.Lists.Add(listinstance);
                ctx.Web.ApplyProvisioningTemplate(template);


                var rowCount = ctx.Web.GetListByTitle(datarowListName).ItemCount;
                Assert.IsTrue(rowCount == 3, "Row count not equals 3");

                listinstance.DataRows.KeyColumn = "Key";
                listinstance.DataRows.UpdateBehavior = UpdateBehavior.Skip;
                ctx.Web.ApplyProvisioningTemplate(template);

                rowCount = ctx.Web.GetListByTitle(datarowListName).ItemCount;
                Assert.IsTrue(rowCount == 3, "Row count not equals 3");

                listinstance.DataRows.UpdateBehavior = UpdateBehavior.Overwrite;
                ctx.Web.ApplyProvisioningTemplate(template);

                rowCount = ctx.Web.GetListByTitle(datarowListName).ItemCount;
                Assert.IsTrue(rowCount == 3, "Row count not equals 3");

                listinstance.DataRows.Add(new DataRow(new Dictionary<string, string> { { "Title", "Test -4-" }, { "Key", "4" } }, "4"));
                ctx.Web.ApplyProvisioningTemplate(template);

                rowCount = ctx.Web.GetListByTitle(datarowListName).ItemCount;
                Assert.IsTrue(rowCount == 4, "Row count not equals 4");

            }
        }

        [TestMethod]
        public void CanUpdateDefaultContentTypeWithoutModifyingContentTypeNewButtonVisibility()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                // create content types
                var documentCtype = web.ContentTypes.GetById(BuiltInContentTypeId.Document);
                var newCtypeInfo1 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType1",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };
                var newCtypeInfo2 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType2",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };
                var newCtypeInfo3 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType3",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };
                var newCtypeInfo4 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType4",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };
                var newCtypeInfo5 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType5",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };

                var newCtype1 = web.ContentTypes.Add(newCtypeInfo1);
                var newCtype2 = web.ContentTypes.Add(newCtypeInfo2);
                var newCtype3 = web.ContentTypes.Add(newCtypeInfo3);
                var newCtype4 = web.ContentTypes.Add(newCtypeInfo4);
                var newCtype5 = web.ContentTypes.Add(newCtypeInfo5);
                clientContext.Load(newCtype1);
                clientContext.Load(newCtype2);
                clientContext.Load(newCtype3);
                clientContext.Load(newCtype4);
                clientContext.Load(newCtype5);
                clientContext.ExecuteQueryRetry();

                var newList = new ListCreationInformation()
                {
                    TemplateType = (int)ListTemplateType.DocumentLibrary,
                    Title = listName,
                    Url = listName
                };

                var doclib = clientContext.Web.Lists.Add(newList);
                doclib.ContentTypesEnabled = true;
                doclib.ContentTypes.AddExistingContentType(newCtype1);
                doclib.ContentTypes.AddExistingContentType(newCtype3);
                doclib.ContentTypes.AddExistingContentType(newCtype4);
                doclib.Update();

                clientContext.Load(newCtype1, ct => ct.Id);
                clientContext.Load(newCtype2, ct => ct.Id);
                clientContext.Load(newCtype3, ct => ct.Id);
                clientContext.Load(newCtype4, ct => ct.Id);
                clientContext.Load(newCtype5, ct => ct.Id);

                clientContext.Load(doclib.ContentTypes);
                clientContext.Load(doclib.RootFolder, rf => rf.ContentTypeOrder);
                clientContext.ExecuteQueryRetry();

                var contentTypeOrder = doclib.RootFolder.ContentTypeOrder;
                //Make a content type hidden in the new button.
                contentTypeOrder.Remove(contentTypeOrder.First(ct => ct.GetParentIdValue().Equals(newCtype3.Id.StringValue, StringComparison.OrdinalIgnoreCase)));

                doclib.RootFolder.UniqueContentTypeOrder = contentTypeOrder;
                Assert.IsTrue(contentTypeOrder.ElementAt(0).GetParentIdValue().Equals(BuiltInContentTypeId.Document, StringComparison.OrdinalIgnoreCase));

                doclib.RootFolder.Update();

                clientContext.ExecuteQueryRetry();

                var template = new ProvisioningTemplate();
                var listInstance = new Core.Framework.Provisioning.Model.ListInstance();

                listInstance.Url = listName;
                listInstance.Title = listName;
                listInstance.TemplateType = (int)ListTemplateType.DocumentLibrary;
                listInstance.ContentTypesEnabled = true;
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding() { ContentTypeId = newCtype1.Id.StringValue, Default = true });
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding() { ContentTypeId = newCtype2.Id.StringValue, Hidden = false });
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding() { ContentTypeId = newCtype4.Id.StringValue, Remove = true });
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding() { ContentTypeId = newCtype5.Id.StringValue });

                template.Lists.Add(listInstance);

                var parser = new TokenParser(clientContext.Web, template);

                // Update the List with new default content type
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(clientContext.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(clientContext.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = clientContext.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                clientContext.Load(doclib.RootFolder, rf => rf.UniqueContentTypeOrder);
                clientContext.ExecuteQueryRetry();

                var actualContentTypeOrder = doclib.RootFolder.UniqueContentTypeOrder;
                bool isContentType2VisibleInNewButton = actualContentTypeOrder.FirstOrDefault(ct => ct.GetParentIdValue().Equals(newCtype2.Id.StringValue, StringComparison.OrdinalIgnoreCase)) != null;
                bool isContentType3VisibleInNewButton = actualContentTypeOrder.FirstOrDefault(ct => ct.GetParentIdValue().Equals(newCtype3.Id.StringValue, StringComparison.OrdinalIgnoreCase)) != null;
                bool isContentType4VisibleInNewButton = actualContentTypeOrder.FirstOrDefault(ct => ct.GetParentIdValue().Equals(newCtype4.Id.StringValue, StringComparison.OrdinalIgnoreCase)) != null;
                bool isContentType5VisibleInNewButton = actualContentTypeOrder.FirstOrDefault(ct => ct.GetParentIdValue().Equals(newCtype5.Id.StringValue, StringComparison.OrdinalIgnoreCase)) != null;

                bool contentType4ExistsInList = doclib.ContentTypeExistsById(newCtype4.Id.StringValue);

                Assert.IsTrue(isContentType2VisibleInNewButton, "Content type 2 has not been made visible in the new button");
                Assert.IsFalse(isContentType3VisibleInNewButton, "Content type 3 has incorrectly been made visible in the new button");
                Assert.IsFalse(isContentType4VisibleInNewButton, "Content type 4 has incorrectly been made visible in the new button");
                Assert.IsTrue(isContentType5VisibleInNewButton, "Content type 5 has not been made visible in the new button");
                Assert.IsFalse(contentType4ExistsInList, "Content type 4 has not been removed from the list content types");
            }
        }
        [TestMethod]
        public void CanRemoveContentTypeWithoutModifyingContentTypeNewButtonVisibility()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                // create content types
                var documentCtype = web.ContentTypes.GetById(BuiltInContentTypeId.Document);
                var newCtypeInfo1 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType1",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };
                var newCtypeInfo2 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType2",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };
                var newCtypeInfo3 = new ContentTypeCreationInformation()
                {
                    Name = "Test_ContentType3",
                    ParentContentType = documentCtype,
                    Group = "Test content types",
                    Description = "This is a test content type"
                };

                var newCtype1 = web.ContentTypes.Add(newCtypeInfo1);
                var newCtype2 = web.ContentTypes.Add(newCtypeInfo2);
                var newCtype3 = web.ContentTypes.Add(newCtypeInfo3);
                clientContext.Load(newCtype1);
                clientContext.Load(newCtype2);
                clientContext.Load(newCtype3);
                clientContext.ExecuteQueryRetry();

                var newList = new ListCreationInformation()
                {
                    TemplateType = (int)ListTemplateType.DocumentLibrary,
                    Title = listName,
                    Url = listName
                };

                var doclib = clientContext.Web.Lists.Add(newList);
                doclib.ContentTypesEnabled = true;
                doclib.ContentTypes.AddExistingContentType(newCtype1);
                doclib.ContentTypes.AddExistingContentType(newCtype3);
                doclib.Update();

                clientContext.Load(newCtype1, ct => ct.Id);
                clientContext.Load(newCtype2, ct => ct.Id);
                clientContext.Load(newCtype3, ct => ct.Id);

                clientContext.Load(doclib.ContentTypes);
                clientContext.Load(doclib.RootFolder, rf => rf.ContentTypeOrder);
                clientContext.ExecuteQueryRetry();

                var contentTypeOrder = doclib.RootFolder.ContentTypeOrder;
                //Make a content type hidden in the new button.
                contentTypeOrder.Remove(contentTypeOrder.First(ct => ct.GetParentIdValue().Equals(newCtype3.Id.StringValue, StringComparison.OrdinalIgnoreCase)));

                doclib.RootFolder.UniqueContentTypeOrder = contentTypeOrder;
                Assert.IsTrue(contentTypeOrder.ElementAt(0).GetParentIdValue().Equals(BuiltInContentTypeId.Document, StringComparison.OrdinalIgnoreCase));

                doclib.RootFolder.Update();
                doclib.Update();

                clientContext.ExecuteQueryRetry();

                var template = new ProvisioningTemplate();
                var listInstance = new Core.Framework.Provisioning.Model.ListInstance();

                listInstance.Url = listName;
                listInstance.Title = listName;
                listInstance.TemplateType = (int)ListTemplateType.DocumentLibrary;
                listInstance.ContentTypesEnabled = true;
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding() { ContentTypeId = newCtype1.Id.StringValue, Default = true });
                listInstance.ContentTypeBindings.Add(new ContentTypeBinding() { ContentTypeId = newCtype2.Id.StringValue, Hidden = false });

                template.Lists.Add(listInstance);

                var parser = new TokenParser(clientContext.Web, template);

                // Update the List with new default content type
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(clientContext.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(clientContext.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = clientContext.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                clientContext.Load(doclib.RootFolder, rf => rf.UniqueContentTypeOrder);
                clientContext.ExecuteQueryRetry();

                var actualContentTypeOrder = doclib.RootFolder.UniqueContentTypeOrder;
                bool isHiddenContentTypeStillHidden = actualContentTypeOrder.FirstOrDefault(ct => ct.GetParentIdValue().Equals(newCtype3.Id.StringValue, StringComparison.OrdinalIgnoreCase)) == null;
                bool isContentType2VisibleInNewButton = actualContentTypeOrder.FirstOrDefault(ct => ct.GetParentIdValue().Equals(newCtype2.Id.StringValue, StringComparison.OrdinalIgnoreCase)) != null;

                Assert.IsTrue(isHiddenContentTypeStillHidden, "Content type has incorrectly been made visible in the new button");
                Assert.IsTrue(isContentType2VisibleInNewButton, "Content type 2 has not been made visible in the new button");
            }
        }
    }
}