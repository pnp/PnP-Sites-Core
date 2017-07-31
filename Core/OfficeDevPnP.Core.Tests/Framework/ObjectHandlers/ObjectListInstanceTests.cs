using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Xml.Linq;

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

        [TestInitialize]
        public void Initialize()
        {
            listName = string.Format("Test_{0}", DateTime.Now.Ticks);

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
    }
}
