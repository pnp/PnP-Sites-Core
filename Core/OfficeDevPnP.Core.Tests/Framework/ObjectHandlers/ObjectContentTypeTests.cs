using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using ContentType = OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType;
using Field = OfficeDevPnP.Core.Framework.Provisioning.Model.Field;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectContentTypeTests
    {
        private const string ElementSchema = @"<ContentType ID=""0x010100503B9E20E5455344BFAC2292DC6FE805"" Name=""Test Content Type"" Group=""PnP"" Version=""1"" xmlns=""http://schemas.microsoft.com/sharepoint/v3"" />";

        private const string SubwebUrl = "9e92eac775b64d31a0fb771de86983aa";

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var ct = ctx.Web.GetContentTypeByName("Test Content Type");
                //var field = ctx.Web.Fields.GetByInternalNameOrTitle("TestField");
                if (ct != null)
                {
                    ct.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
                try
                {
                    ctx.Web.RemoveFieldByInternalName("PnPTestField");
                }
                catch
                {
                }
                var subwebs = ctx.LoadQuery(ctx.Web.Webs);
                ctx.ExecuteQueryRetry();

                if (subwebs.FirstOrDefault(w => w.Url.EndsWith(SubwebUrl)) != null)
                {
                    subwebs.FirstOrDefault(w => w.Url.EndsWith(SubwebUrl)).DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();


            var contentType = new ContentType()
            {
                Id = "0x010100503B9E20E5455344BFAC2292DC6FE805",
                Name = "Test Content Type",
                Group = "PnP",
                Description = "Test Description",
                Overwrite = true,
                Hidden = false
            };

            contentType.FieldRefs.Add(new FieldRef()
            {
                Id = BuiltInFieldId.Category,
                DisplayName = "Test Category",
            });
            template.ContentTypes.Add(contentType);

            using (var ctx = TestCommon.CreateClientContext())
            {
                TokenParser parser = new TokenParser(ctx.Web, template);
                new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var ct = ctx.Web.GetContentTypeByName("Test Content Type");

                Assert.IsNotNull(ct);

            }
        }

        [TestMethod]
        public void CanProvisionToObjectsToSubweb()
        {
            var template = new ProvisioningTemplate();


            var contentType = new ContentType()
            {
                Id = "0x010100503B9E20E5455344BFAC2292DC6FE805",
                Name = "Test Content Type",
                Group = "PnP",
                Description = "Test Description",
                Overwrite = true,
                Hidden = false
            };

            template.ContentTypes.Add(contentType);

            using (var ctx = TestCommon.CreateClientContext())
            {
                // Create subweb
                var web = ctx.Web.Webs.Add(new WebCreationInformation()
                {
                    Description = "Subweb",
                    Language = 1033,
                    Title = "Subweb",
                    Url = SubwebUrl,
                    WebTemplate = "STS#0"
                });
                ctx.Load(web);
                ctx.ExecuteQueryRetry();

                TokenParser parser = new TokenParser(web, template);

                var applyingInformation = new ProvisioningTemplateApplyingInformation();

                new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(web, template, parser, applyingInformation);

                var ct = web.GetContentTypeByName("Test Content Type");

                Assert.IsNull(ct);

                applyingInformation.ProvisionContentTypesToSubWebs = true;

                new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(web, template, parser, applyingInformation);

                ct = web.GetContentTypeByName("Test Content Type");

                Assert.IsNotNull(ct);
            }
        }

        [TestMethod]
        public void FieldUsingTokensAreCorrectlyOrdered()
        {
            var template = new ProvisioningTemplate();

            template.Parameters.Add("TestFieldPrefix", "PnP");


            var contentType = new ContentType
            {
                Id = "0x010100503B9E20E5455344BFAC2292DC6FE805",
                Name = "Test Content Type",
                Group = "PnP",
                Description = "Test Description",
                Overwrite = true,
                Hidden = false
            };

            var nonOobField = new Field
            {
                SchemaXml = "<Field ID=\"{dd6b7dae-1281-458d-a66c-01b0c7b7930b}\" Name=\"{parameter:TestFieldPrefix}TestField\" DisplayName=\"TestField\" Type=\"Note\" Group=\"PnP\" Description=\"\" />"
            };
            template.SiteFields.Add(nonOobField);

            contentType.FieldRefs.Add(new FieldRef("{parameter:TestFieldPrefix}TestField")
            {
                Id = new Guid("{dd6b7dae-1281-458d-a66c-01b0c7b7930b}")
            });

            contentType.FieldRefs.Add(new FieldRef("AssignedTo")
            {
                Id = BuiltInFieldId.AssignedTo
            });
            template.ContentTypes.Add(contentType);

            using (var ctx = TestCommon.CreateClientContext())
            {
                TokenParser parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                var ct = ctx.Web.GetContentTypeByName("Test Content Type");
                ct.EnsureProperty(x => x.FieldLinks);
                Assert.AreEqual(ct.FieldLinks[0].Id, template.ContentTypes.First().FieldRefs[0].Id);
                Assert.AreEqual(ct.FieldLinks[1].Id, template.ContentTypes.First().FieldRefs[1].Id);

            }


        }

        [TestMethod]
        public void CanCreateEntities()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                // Provision a test content type
                var ct = ctx.Web.GetContentTypeByName("Test Content Type");
                if (ct == null)
                {
                    var provisionTemplate = new ProvisioningTemplate();
                    var contentType = new ContentType()
                    {
                        Id = "0x010100503B9E20E5455344BFAC2292DC6FE805",
                        Name = "Test Content Type",
                        Group = "PnP",
                        Description = "Test Description",
                        Overwrite = true,
                        Hidden = false
                    };

                    contentType.FieldRefs.Add(new FieldRef()
                    {
                        Id = BuiltInFieldId.Category,
                        DisplayName = "Test Category",
                    });
                    provisionTemplate.ContentTypes.Add(contentType);
                    TokenParser parser = new TokenParser(ctx.Web, provisionTemplate);
                    new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, provisionTemplate, parser, new ProvisioningTemplateApplyingInformation());
                }

                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template = new ProvisioningTemplate();
                template = new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.ContentTypes.Any());
                Assert.IsInstanceOfType(template.ContentTypes, typeof(Core.Framework.Provisioning.Model.ContentTypeCollection));
            }
        }

        [TestMethod]
        public void WorkflowTaskOutcomeFieldIsUnique()
        {
            var template = new ProvisioningTemplate();

            var contentType = new ContentType
            {
                Id = "0x0108003365C4474CAE8C42BCE396314E88E51F008E5B850C364947248508D252250ED723",
                Name = "Test Custom Outcome Workflow Task",
                Group = "PnP",
                Description = "Ensure inherited workflow task displays correct custom OutcomeChoice",
                Overwrite = true,
                Hidden = false
            };

            var nonOobField = new Field
            {
                SchemaXml = "<Field ID=\"{35e4bd1f-c1a3-4bf2-bf86-4470c2e8bcfd}\" Type=\"OutcomeChoice\" StaticName=\"AuthorReviewOutcome\" Name=\"AuthorReviewOutcome\" DisplayName=\"AuthorReviewOutcome\" Group=\"PnP\">"
                      + "<Default>Approved</Default>"
                      + "<CHOICES>"
                      + "<CHOICE>Approved</CHOICE>"
                      + "<CHOICE>Rejected</CHOICE>"
                      + "<CHOICE>Reassign</CHOICE>"
                      + "</CHOICES>"
                    + "</Field>"
            };
            template.SiteFields.Add(nonOobField);

            contentType.FieldRefs.Add(new FieldRef("AuthorReviewOutcome")
            {
                Id = new Guid("{35e4bd1f-c1a3-4bf2-bf86-4470c2e8bcfd}")
            });

            template.ContentTypes.Add(contentType);

            using (var ctx = TestCommon.CreateClientContext())
            {
                TokenParser parser = new TokenParser(ctx.Web, template);
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                var ct = ctx.Web.GetContentTypeByName("Test Custom Outcome Workflow Task");
                ct.EnsureProperty(x => x.Fields);
                Assert.AreEqual(ct.Fields.Count(f => f.FieldTypeKind == FieldType.OutcomeChoice), 1);

            }
        }
    }
}
