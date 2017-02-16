using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class ContentTypeImplementation : ImplementationBase
    {
        internal void SiteCollectionContentTypeAdding(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {

                // Add supporting files, note that files validation will be done in the files test cases
                TestProvisioningTemplate(cc, "contenttype_files.xml", Handlers.Files);

                // Ensure we can test clean
                DeleteContentTypes(cc);

                // Add content types
                var result = TestProvisioningTemplate(cc, "contenttype_add.xml", Handlers.ContentTypes | Handlers.Fields);
                ContentTypeValidator cv = new ContentTypeValidator(cc.Web);
                Assert.IsTrue(cv.Validate(result.SourceTemplate.ContentTypes, result.TargetTemplate.ContentTypes, result.TargetTokenParser));

                // change content types
                var result2 = TestProvisioningTemplate(cc, "contenttype_delta_1.xml", Handlers.ContentTypes);
                Assert.IsTrue(cv.Validate(result2.SourceTemplate.ContentTypes, result2.TargetTemplate.ContentTypes, result2.TargetTokenParser));
            }
        }

        #region Helper methods
        private void DeleteContentTypes(ClientContext cc)
        {
            // Drop the content types
            cc.Load(cc.Web.ContentTypes, f => f.Include(t => t.Name));
            cc.ExecuteQueryRetry();

            foreach (var ct in cc.Web.ContentTypes.ToList())
            {
                if (ct.Name.StartsWith("CT_"))
                {
                    ct.DeleteObject();
                }
            }
            cc.ExecuteQueryRetry();

            // Drop the fields
            DeleteFields(cc);
        }

        private void DeleteFields(ClientContext cc)
        {
            cc.Load(cc.Web.Fields, f => f.Include(t => t.InternalName));
            cc.ExecuteQueryRetry();

            foreach (var field in cc.Web.Fields.ToList())
            {
                // First drop the fields that have 2 _'s...convention used to name the fields dependent on a lookup.
                if (field.InternalName.Replace("FLD_CT_", "").IndexOf("_") > 0)
                {
                    if (field.InternalName.StartsWith("FLD_CT_"))
                    {
                        field.DeleteObject();
                    }
                }
            }

            foreach (var field in cc.Web.Fields.ToList())
            {
                if (field.InternalName.StartsWith("FLD_CT_"))
                {
                    field.DeleteObject();
                }
            }

            cc.ExecuteQueryRetry();

        }
        #endregion

    }
}