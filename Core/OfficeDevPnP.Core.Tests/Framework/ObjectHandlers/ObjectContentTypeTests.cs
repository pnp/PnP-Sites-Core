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
                ctx.Web.RemoveFieldByInternalName("PnPTestField");

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
                new ObjectContentType().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var ct = ctx.Web.GetContentTypeByName("Test Content Type");

                Assert.IsNotNull(ct);

            }
        }

        [TestMethod]
        public void FieldUsingTokensAreCorrectlyOrdered()
        {
            var template = new ProvisioningTemplate();

            template.Parameters.Add("TestFieldPrefix","PnP");
            

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
                Id= new Guid("{dd6b7dae-1281-458d-a66c-01b0c7b7930b}")
            });

            contentType.FieldRefs.Add(new FieldRef("AssignedTo")
            {
                Id = BuiltInFieldId.AssignedTo
            });
            template.ContentTypes.Add(contentType);

            using (var ctx = TestCommon.CreateClientContext())
            {
                TokenParser parser = new TokenParser(ctx.Web, template);
                new ObjectField().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
                new ObjectContentType().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
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
                    new ObjectContentType().ProvisionObjects(ctx.Web, provisionTemplate, parser, new ProvisioningTemplateApplyingInformation());
                }

                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template = new ProvisioningTemplate();
                template = new ObjectContentType().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.ContentTypes.Any());
                Assert.IsInstanceOfType(template.ContentTypes, typeof(Core.Framework.Provisioning.Model.ContentTypeCollection));
            }
        }
    }
}
