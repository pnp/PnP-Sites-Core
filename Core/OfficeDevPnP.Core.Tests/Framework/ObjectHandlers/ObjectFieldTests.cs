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
    public class ObjectFieldTests
    {
        private const string ElementSchema = @"<Field xmlns=""http://schemas.microsoft.com/sharepoint/v3"" StaticName=""DemoField"" DisplayName=""Test Field"" Type=""Text"" ID=""{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}"" Group=""PnP"" Required=""true""/>";
        private Guid fieldId = Guid.Parse("{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}");

        private const string CalculatedFieldElementSchema = @"<Field Name=""CalculatedField"" StaticName=""CalculatedField"" DisplayName=""Test Calculated Field"" Type=""Calculated"" ResultType=""Text"" ID=""{D1A33456-9FEB-4D8E-AFFA-177EACCE4B70}"" Group=""PnP"" ReadOnly=""TRUE"" ><Formula>=DemoField&amp;""DemoField""</Formula><FieldRefs><FieldRef Name=""DemoField"" ID=""{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}"" /></FieldRefs></Field>";
        private const string TokenizedCalculatedFieldElementSchema = @"<Field Name=""CalculatedField"" StaticName=""CalculatedField"" DisplayName=""Test Calculated Field"" Type=""Calculated"" ResultType=""Text"" ID=""{D1A33456-9FEB-4D8E-AFFA-177EACCE4B70}"" Group=""PnP"" ReadOnly=""TRUE"" ><Formula>=[{fieldtitle:DemoField}]&amp;""DemoField""</Formula></Field>";
        private Guid calculatedFieldId = Guid.Parse("{D1A33456-9FEB-4D8E-AFFA-177EACCE4B70}");

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var f = ctx.Web.GetFieldById<FieldText>(fieldId); // Guid matches ID in field caml.
                var calculatedField = ctx.Web.GetFieldById<FieldCalculated>(calculatedFieldId); // Guid matches ID in field caml.

                bool fieldDeleted = false;
                if (f != null)
                {
                    f.DeleteObject();
                    fieldDeleted = true;
                }
                if (calculatedField != null)
                {
                    calculatedField.DeleteObject();
                    fieldDeleted = true;
                }
                if (fieldDeleted)
                {
                    ctx.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var f = ctx.Web.GetFieldById<FieldText>(fieldId);

                Assert.IsNotNull(f);
                Assert.IsInstanceOfType(f, typeof(FieldText));
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
                template = new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.SiteFields.Any());
                Assert.IsInstanceOfType(template.SiteFields, typeof(Core.Framework.Provisioning.Model.FieldCollection));
            }
        }

        [TestMethod]
        public void CanProvisionCalculatedField()
        {
            var template = new ProvisioningTemplate();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = TokenizedCalculatedFieldElementSchema });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var f = ctx.Web.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(f);
                Assert.IsInstanceOfType(f, typeof(FieldCalculated));
                Assert.IsFalse(f.Formula.Contains('#') || f.Formula.Contains('?'), "Calculated field was not provisioned properly");
            }
        }

        [TestMethod]
        public void CanUpdateCalculatedField()
        {
            var template = new ProvisioningTemplate();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = TokenizedCalculatedFieldElementSchema });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var f1 = ctx.Web.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(f1);
                Assert.IsInstanceOfType(f1, typeof(FieldCalculated));
                Assert.IsFalse(f1.Formula.Contains('#') || f1.Formula.Contains('?'), "Calculated field was not provisioned properly the first time");

                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var f2 = ctx.Web.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(f2);
                Assert.IsInstanceOfType(f2, typeof(FieldCalculated));
                Assert.IsFalse(f2.Formula.Contains('#') || f2.Formula.Contains('?'), "Calculated field was not provisioned properly the second time");
            }
        }

        [TestMethod]
        public void CanExtractCalculatedFieldsLast()
        {
            var template = new ProvisioningTemplate();
            var extractedTemplate = new ProvisioningTemplate();
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = TokenizedCalculatedFieldElementSchema });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var provisioningTemplateCreationInformation = new ProvisioningTemplateCreationInformation(ctx.Web);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ExtractObjects(ctx.Web, extractedTemplate, provisioningTemplateCreationInformation);

                XElement fieldElement = XElement.Parse(extractedTemplate.SiteFields.Last().SchemaXml);

                Assert.AreEqual(calculatedFieldId, Guid.Parse(fieldElement.Attribute("ID").Value), "Calculated field is not extracted as last field");
            }
        }

        [TestMethod]
        public void CanExtractCalculatedField()
        {
            var template = new ProvisioningTemplate();
            var listInstance = new ListInstance();
            var extractedTemplate = new ProvisioningTemplate();

            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });
            template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = TokenizedCalculatedFieldElementSchema });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var f = ctx.Web.GetFieldById<FieldCalculated>(calculatedFieldId);

                Assert.IsNotNull(f);
                Assert.IsInstanceOfType(f, typeof(FieldCalculated));
                Assert.IsFalse(f.Formula.Contains('#') || f.Formula.Contains('?'), "Calculated field was not provisioned properly");

                var provisioningTemplateCreationInformation = new ProvisioningTemplateCreationInformation(ctx.Web);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ExtractObjects(ctx.Web, extractedTemplate, provisioningTemplateCreationInformation);

                XElement fieldElement = XElement.Parse(extractedTemplate.SiteFields.First(cf => Guid.Parse(XElement.Parse(cf.SchemaXml).Attribute("ID").Value).Equals(calculatedFieldId)).SchemaXml);
                var formula = fieldElement.Descendants("Formula").FirstOrDefault();

                Assert.AreEqual(@"=[{fieldtitle:DemoField}]&""DemoField""", formula.Value, true, "Calculated field formula is not extracted properly");
            }
        }
    }
}
