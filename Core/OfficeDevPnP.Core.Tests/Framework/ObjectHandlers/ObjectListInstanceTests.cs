using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Xml;
using System.Text;
using System.IO;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectListInstanceTests
    {
        private const string ElementSchema = @"<Field xmlns=""http://schemas.microsoft.com/sharepoint/v3"" Name=""DemoField"" StaticName=""DemoField"" DisplayName=""Test Field"" Type=""Text"" ID=""{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}"" Group=""PnP"" Required=""true""/>";
        private Guid fieldId = Guid.Parse("{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}");

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
            var template = new ProvisioningTemplate();
            var listInstance = new Core.Framework.Provisioning.Model.ListInstance();

            listInstance.Url = string.Format("lists/{0}", listName);
            listInstance.Title = listName;
            listInstance.TemplateType = (int)ListTemplateType.GenericList;

            Dictionary<string, string> dataValues = new Dictionary<string, string>();
            dataValues.Add("Title", "Test");
            DataRow dataRow = new DataRow(dataValues);

            listInstance.DataRows.Add(dataRow);

            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);

                // Create the List
                parser = new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                // Load DataRows
                new ObjectListInstanceDataRows().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var items = list.GetItems(CamlQuery.CreateAllItemsQuery());
                ctx.Load(items, itms => itms.Include(item => item["Title"]));
                ctx.ExecuteQueryRetry();

                Assert.IsTrue(items.Count == 1);
                Assert.IsTrue(items[0]["Title"].ToString() == "Test");
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
                template = new ObjectListInstance().ExtractObjects(ctx.Web, template, creationInfo);

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
                new ObjectField().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

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
                new ObjectField().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f1 = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f1);
                Assert.IsInstanceOfType(f1, typeof(FieldCalculated));
                Assert.IsFalse(f1.Formula.Contains('#') || f1.Formula.Contains('?'), "Calculated field was not provisioned properly the first time");

                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

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
                new ObjectField().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

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
                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

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
                new ObjectField().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var list = ctx.Web.GetListByUrl(listInstance.Url);
                Assert.IsNotNull(list);

                var rf = list.GetFieldById<FieldText>(fieldId);
                var f1 = list.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(rf, "Referenced field not added");
                Assert.IsNotNull(f1);
                Assert.IsInstanceOfType(f1, typeof(FieldCalculated));
                Assert.IsFalse(f1.Formula.Contains('#') || f1.Formula.Contains('?'), "Calculated field was not provisioned properly the first time");

                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

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
                new ObjectField().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

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
                new ObjectListInstance().ExtractObjects(ctx.Web, extractedTemplate, provisioningTemplateCreationInformation);

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
        public void ChangeOOBFieldTitles_EnUs()
        {
            CheckOOBFieldsAreHidden(1033);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_ArSa()
        {
            CheckOOBFieldsAreHidden(1025);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_NlNl()
        {
            CheckOOBFieldsAreHidden(1043);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_AzLatn()
        {
            CheckOOBFieldsAreHidden(1068);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_EuEs()
        {
            CheckOOBFieldsAreHidden(1069);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_BsLatnBa()
        {
            CheckOOBFieldsAreHidden(5146);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_BgBg()
        {
            CheckOOBFieldsAreHidden(1026);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_CaEs()
        {
            CheckOOBFieldsAreHidden(1027);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_ZhCn()
        {
            CheckOOBFieldsAreHidden(2052);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_ZhTw()
        {
            CheckOOBFieldsAreHidden(1028);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_HrHr()
        {
            CheckOOBFieldsAreHidden(1050);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_CsCz()
        {
            CheckOOBFieldsAreHidden(1029);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_DaDk()
        {
            CheckOOBFieldsAreHidden(1030);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_PrsAf()
        {
            CheckOOBFieldsAreHidden(1164);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_EtEe()
        {
            CheckOOBFieldsAreHidden(1061);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_FiFi()
        {
            CheckOOBFieldsAreHidden(1035);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_FrFr()
        {
            CheckOOBFieldsAreHidden(1036);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_GlEs()
        {
            CheckOOBFieldsAreHidden(1110);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_DeDe()
        {
            CheckOOBFieldsAreHidden(1031);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_ElGr()
        {
            CheckOOBFieldsAreHidden(1032);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_HeIl()
        {
            CheckOOBFieldsAreHidden(1037);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_HiIn()
        {
            CheckOOBFieldsAreHidden(1081);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_HuHu()
        {
            CheckOOBFieldsAreHidden(1038);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_IdId()
        {
            CheckOOBFieldsAreHidden(1057);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_GaIe()
        {
            CheckOOBFieldsAreHidden(2108);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_ItIt()
        {
            CheckOOBFieldsAreHidden(1040);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_JaJp()
        {
            CheckOOBFieldsAreHidden(1041);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_KkKz()
        {
            CheckOOBFieldsAreHidden(1087);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_KoKr()
        {
            CheckOOBFieldsAreHidden(1042);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_LvLv()
        {
            CheckOOBFieldsAreHidden(1062);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_LtLt()
        {
            CheckOOBFieldsAreHidden(1063);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_MkMk()
        {
            CheckOOBFieldsAreHidden(1071);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_MsMy()
        {
            CheckOOBFieldsAreHidden(1086);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_NbNo()
        {
            CheckOOBFieldsAreHidden(1044);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_PlPl()
        {
            CheckOOBFieldsAreHidden(1045);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_PtBr()
        {
            CheckOOBFieldsAreHidden(1046);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_PtPt()
        {
            CheckOOBFieldsAreHidden(2070);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_RoRo()
        {
            CheckOOBFieldsAreHidden(1048);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_RuRu()
        {
            CheckOOBFieldsAreHidden(1049);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_SrCyrlRs()
        {
            CheckOOBFieldsAreHidden(10266);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_SrLatnRs()
        {
            CheckOOBFieldsAreHidden(9242);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_SkSk()
        {
            CheckOOBFieldsAreHidden(1051);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_SlSi()
        {
            CheckOOBFieldsAreHidden(1060);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_EsEs()
        {
            CheckOOBFieldsAreHidden(3082);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_SvSe()
        {
            CheckOOBFieldsAreHidden(1053);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_ThTh()
        {
            CheckOOBFieldsAreHidden(1054);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_TrTr()
        {
            CheckOOBFieldsAreHidden(1055);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_UkUa()
        {
            CheckOOBFieldsAreHidden(1058);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_ViVn()
        {
            CheckOOBFieldsAreHidden(1066);
        }

        [TestMethod]
        public void ChangeOOBFieldTitles_CyGb()
        {
            CheckOOBFieldsAreHidden(1106);
        }

        #region HelperMethods

        public void CheckOOBFieldsAreHidden(int localeId)
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var siteTitle = TestCommon.GetMultiLingualSubSiteTitle(localeId);
                Web subWeb = ctx.Web.GetWeb(siteTitle) ?? ctx.Web.CreateWeb(new Entities.SiteEntity() { Lcid = (uint)localeId, Title = siteTitle, Url = siteTitle });
                var cultureInfo = CultureInfo.GetCultureInfo(localeId);

                CheckOOBFieldsAreHiddenForList(subWeb, cultureInfo);
                CheckOOBFieldsAreHiddenForLibrary(subWeb, cultureInfo);
                CheckOOBFieldsAreHiddenForImagesLibrary(subWeb, cultureInfo);
            }
        }

        //Check if the default fields are hidden (return 0 fields), change title of 'Title' field, should be returned (return 3 fields)
        public void CheckOOBFieldsAreHiddenForList(Web subWeb, CultureInfo cultureInfo)
        {
            string listTitle = "CustomList";

            //Delete list if it exists
            var list = subWeb.GetListByTitle(listTitle);
            if (list != null)
            {
                list.DeleteObject();
                list.Context.ExecuteQueryRetry();
            }

            //Create new list (to ensure, no custom fields are present)
            list = subWeb.CreateList(ListTemplateType.GenericList, listTitle, false, true);

            //Extract provisioning template of the list. No fields should be present. (Because there are only OOB field at this moment)
            var template = subWeb.GetProvisioningTemplate(new ProvisioningTemplateCreationInformation(subWeb)
            {
                HandlersToProcess = Handlers.Lists
            });
            Assert.IsTrue(template.Lists.Any());
            var templateList = template.Lists.SingleOrDefault(l => l.Title.Equals(listTitle));
            Assert.IsNotNull(templateList);

            Assert.AreEqual(0, templateList.FieldRefs.Count);

            //Set title's display name to custom value      
            ChangeListFieldTitle(list, "Title", "CustomTitle", cultureInfo);
            ChangeListFieldTitle(list, "LinkTitle", "CustomTitle", cultureInfo);
            ChangeListFieldTitle(list, "LinkTitleNoMenu", "CustomTitle", cultureInfo);

            //Extract provisioning template of the list. Title field should be present
            template = subWeb.GetProvisioningTemplate(new ProvisioningTemplateCreationInformation(subWeb)
            {
                HandlersToProcess = Handlers.Lists
            });
            Assert.IsTrue(template.Lists.Any());
            templateList = template.Lists.SingleOrDefault(l => l.Title.Equals(listTitle));
            Assert.IsNotNull(templateList);
            Assert.AreEqual(3, templateList.FieldRefs.Count);
            Assert.IsTrue(templateList.FieldRefs.Any(f => f.Name.Equals("Title")));
            Assert.IsTrue(templateList.FieldRefs.Any(f => f.Name.Equals("LinkTitleNoMenu")));
            Assert.IsTrue(templateList.FieldRefs.Any(f => f.Name.Equals("LinkTitle")));
        }
        
        //Check if the default fields are hidden (return 5 fields), change title of 'Title' field, should be returned (return 6 fields)
        public void CheckOOBFieldsAreHiddenForLibrary(Web subWeb, CultureInfo cultureInfo)
        {
            string libraryTitle = "CustomLib";

            //Delete library if it exists
            var lib = subWeb.GetListByTitle(libraryTitle);
            if (lib != null)
            {
                lib.DeleteObject();
                lib.Context.ExecuteQueryRetry();
            }

            //Create new library (to ensure, no custom fields are present)
            lib = subWeb.CreateList(ListTemplateType.DocumentLibrary, libraryTitle, false, true);

            //Extract provisioning template of the list. No fields should be present. (Because there are only OOB field at this moment)
            var template = subWeb.GetProvisioningTemplate(new ProvisioningTemplateCreationInformation(subWeb)
            {
                HandlersToProcess = Handlers.Lists
            });
            Assert.IsTrue(template.Lists.Any());
            var templateList = template.Lists.SingleOrDefault(l => l.Title.Equals(libraryTitle));
            Assert.IsNotNull(templateList);
            
            Assert.AreEqual(0, templateList.FieldRefs.Count);

            ChangeListFieldTitle(lib, "Title", "CustomTitle", cultureInfo);

            //Extract provisioning template of the list. Title field should be present
            template = subWeb.GetProvisioningTemplate(new ProvisioningTemplateCreationInformation(subWeb)
            {
                HandlersToProcess = Handlers.Lists
            });
            Assert.IsTrue(template.Lists.Any());
            templateList = template.Lists.SingleOrDefault(l => l.Title.Equals(libraryTitle));
            Assert.IsNotNull(templateList);
            Assert.AreEqual(1, templateList.FieldRefs.Count);
            Assert.IsTrue(templateList.FieldRefs.Any(f => f.Name.Equals("Title")));
        }

        //Check if the default fields are hidden (return 17 fields), change title of 'Title' field, should be returned (return 18 fields)
        public void CheckOOBFieldsAreHiddenForImagesLibrary(Web subWeb, CultureInfo cultureInfo)
        {
            string imageLibTitle = "CustomImages";

            //Delete image library if it exists
            var imageLib = subWeb.GetListByTitle(imageLibTitle);
            if (imageLib != null)
            {
                imageLib.DeleteObject();
                imageLib.Context.ExecuteQueryRetry();
            }

            //Create new library (to ensure, no custom fields are present)
            imageLib = subWeb.CreateList(ListTemplateType.PictureLibrary, imageLibTitle, false, true);

            //Extract provisioning template of the list. No fields should be present. (Because there are only OOB field at this moment)
            var template = subWeb.GetProvisioningTemplate(new ProvisioningTemplateCreationInformation(subWeb)
            {
                HandlersToProcess = Handlers.Lists
            });
            Assert.IsTrue(template.Lists.Any());
            var templateList = template.Lists.SingleOrDefault(l => l.Title.Equals(imageLibTitle));
            Assert.IsNotNull(templateList);
            Assert.AreEqual(0, templateList.FieldRefs.Count);

            //Set title's display name to custom value
            ChangeListFieldTitle(imageLib, "Title", "CustomTitle", cultureInfo);

            //Extract provisioning template of the list. Title field should be present
            template = subWeb.GetProvisioningTemplate(new ProvisioningTemplateCreationInformation(subWeb)
            {
                HandlersToProcess = Handlers.Lists
            });
            Assert.IsTrue(template.Lists.Any());
            templateList = template.Lists.SingleOrDefault(l => l.Title.Equals(imageLibTitle));
            Assert.IsNotNull(templateList);
            Assert.AreEqual(1, templateList.FieldRefs.Count);
            Assert.IsTrue(templateList.FieldRefs.Any(f => f.Name.Equals("Title")));
        }

        public void ChangeListFieldTitle(List list, string internalName, string customTitle, CultureInfo cultureInfo = null)
        {
            var titleField = list.Fields.GetFieldByInternalName(internalName);
            titleField.Title = customTitle;
            titleField.SchemaXml = Regex.Replace(titleField.SchemaXml, "DisplayName=\"[^\"]*\"", "DisplayName=\"" + customTitle + "\"");
            if (cultureInfo != null)
                titleField.TitleResource.SetValueForUICulture(cultureInfo.Name, customTitle);
            titleField.UpdateAndPushChanges(true);
            list.Context.ExecuteQueryRetry();
        }

        #endregion
    }
}