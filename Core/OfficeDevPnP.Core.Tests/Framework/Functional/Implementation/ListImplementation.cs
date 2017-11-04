using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class ListImplementation : ImplementationBase
    {
        internal void SiteCollectionListAdding(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                // Add supporting files needed during add
                TestProvisioningTemplate(cc, "list_supporting_data_1.xml", Handlers.Fields | Handlers.ContentTypes);

                // Add lists
                var result = TestProvisioningTemplate(cc, "list_add.xml", Handlers.Lists);
                ListInstanceValidator lv = new ListInstanceValidator(cc);
                Assert.IsTrue(lv.Validate(result.SourceTemplate.Lists, result.TargetTemplate.Lists, result.TargetTokenParser));

                // Add supporting files needed during delta testing
                TestProvisioningTemplate(cc, "list_supporting_data_2.xml", Handlers.Files);

                // Delta lists
                var result2 = TestProvisioningTemplate(cc, "list_delta_1.xml", Handlers.Lists);
                ListInstanceValidator lv2 = new ListInstanceValidator(cc);
                Assert.IsTrue(lv2.Validate(result2.SourceTemplate.Lists, result2.TargetTemplate.Lists, result2.TargetTokenParser));
            }

        }

        internal void SiteCollection1605ListAdding(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                // Add lists
                var result = TestProvisioningTemplate(cc, "list_add_1605.xml", Handlers.Lists);
                ListInstanceValidator lv = new ListInstanceValidator(cc);
                lv.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(lv.Validate(result.SourceTemplate.Lists, result.TargetTemplate.Lists, result.TargetTokenParser));

                // Delta lists
                var result2 = TestProvisioningTemplate(cc, "list_delta_1605_1.xml", Handlers.Lists);
                ListInstanceValidator lv2 = new ListInstanceValidator(cc);
                lv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(lv2.Validate(result2.SourceTemplate.Lists, result2.TargetTemplate.Lists, result2.TargetTokenParser));
            }

        }

        internal void SiteCollection1705ListAdding(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                // Add lists
                var result = TestProvisioningTemplate(cc, "list_add_1705.xml", Handlers.Lists);
                ListInstanceValidator lv = new ListInstanceValidator(cc)
                {
                    SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05
                };
                Assert.IsTrue(lv.Validate(result.SourceTemplate.Lists, result.TargetTemplate.Lists, result.TargetTokenParser));

                // Delta lists
                var result2 = TestProvisioningTemplate(cc, "list_delta_1705_1.xml", Handlers.Lists);
                ListInstanceValidator lv2 = new ListInstanceValidator(cc);
                lv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05;
                Assert.IsTrue(lv2.Validate(result2.SourceTemplate.Lists, result2.TargetTemplate.Lists, result2.TargetTokenParser));
            }
        }

        internal void WebListAdding(string siteCollectionUrl, string url)
        {
            using (var cc = TestCommon.CreateClientContext(siteCollectionUrl))
            {
                // Add supporting files needed during add
                TestProvisioningTemplate(cc, "list_supporting_data_1.xml", Handlers.Fields | Handlers.ContentTypes);
            }

            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                // Add lists
                var result = TestProvisioningTemplate(cc, "list_add.xml", Handlers.Lists);
                ListInstanceValidator lv = new ListInstanceValidator(cc);
                Assert.IsTrue(lv.Validate(result.SourceTemplate.Lists, result.TargetTemplate.Lists, result.TargetTokenParser));

                // Add supporting files needed during delta testing
                TestProvisioningTemplate(cc, "list_supporting_data_2.xml", Handlers.Files);

                // Delta lists
                var result2 = TestProvisioningTemplate(cc, "list_delta_1.xml", Handlers.Lists);
                ListInstanceValidator lv2 = new ListInstanceValidator(cc);
                Assert.IsTrue(lv2.Validate(result2.SourceTemplate.Lists, result2.TargetTemplate.Lists, result2.TargetTokenParser));
            }

        }

        internal void Web1605ListAdding(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                // Add lists
                var result = TestProvisioningTemplate(cc, "list_add_1605.xml", Handlers.Lists);
                ListInstanceValidator lv = new ListInstanceValidator(cc);
                lv.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(lv.Validate(result.SourceTemplate.Lists, result.TargetTemplate.Lists, result.TargetTokenParser));

                // Delta lists
                var result2 = TestProvisioningTemplate(cc, "list_delta_1605_1.xml", Handlers.Lists);
                ListInstanceValidator lv2 = new ListInstanceValidator(cc);
                lv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(lv2.Validate(result2.SourceTemplate.Lists, result2.TargetTemplate.Lists, result2.TargetTokenParser));
            }
        }

        internal void Web1705ListAdding(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                // Add lists
                var result = TestProvisioningTemplate(cc, "list_add_1705.xml", Handlers.Lists);
                ListInstanceValidator lv = new ListInstanceValidator(cc)
                {
                    SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05
                };
                Assert.IsTrue(lv.Validate(result.SourceTemplate.Lists, result.TargetTemplate.Lists, result.TargetTokenParser));

                // Delta lists
                var result2 = TestProvisioningTemplate(cc, "list_delta_1705_1.xml", Handlers.Lists);
                ListInstanceValidator lv2 = new ListInstanceValidator(cc);
                lv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05;
                Assert.IsTrue(lv2.Validate(result2.SourceTemplate.Lists, result2.TargetTemplate.Lists, result2.TargetTokenParser));
            }
        }

        #region Helper methods
        private void DeleteLists(ClientContext cc)
        {
            DeleteListsImplementation(cc);
        }

        private static void DeleteListsImplementation(ClientContext cc)
        {
            cc.Load(cc.Web.Lists, f => f.Include(t => t.Title));
            cc.ExecuteQueryRetry();

            foreach (var list in cc.Web.Lists.ToList())
            {
                if (list.Title.StartsWith("LI_"))
                {
                    list.DeleteObject();
                }
            }
            cc.ExecuteQueryRetry();
        }

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