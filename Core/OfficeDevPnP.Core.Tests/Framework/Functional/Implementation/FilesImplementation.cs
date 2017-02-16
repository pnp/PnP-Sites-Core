using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class FilesImplementation : ImplementationBase
    {
        internal void SiteCollectionFiles(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                var result = TestProvisioningTemplate(cc, "files_add.xml", Handlers.Files | Handlers.Lists);
                FilesValidator fv = new FilesValidator();
                Assert.IsTrue(fv.Validate(result.SourceTemplate.Files, cc));
            }
        }

        internal void SiteCollectionDirectoryFiles(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                var result = TestProvisioningTemplate(cc, "files_add_1605.xml", Handlers.Files | Handlers.Lists);
                FilesValidator fv = new FilesValidator();
                fv.SchemaVersion = Core.Framework.Provisioning.Providers.Xml.XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(fv.Validate1605(result.SourceTemplate, cc));
            }
        }

        internal void WebFiles(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                var result = TestProvisioningTemplate(cc, "files_add.xml", Handlers.Files | Handlers.Lists);
                FilesValidator fv = new FilesValidator();
                Assert.IsTrue(fv.Validate(result.SourceTemplate.Files, cc));
            }
        }

        internal void WebDirectoryFiles(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                var result = TestProvisioningTemplate(cc, "files_add_1605.xml", Handlers.Files | Handlers.Lists);
                FilesValidator fv = new FilesValidator();
                fv.SchemaVersion = Core.Framework.Provisioning.Providers.Xml.XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(fv.Validate1605(result.SourceTemplate, cc));
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
        #endregion
    }
}