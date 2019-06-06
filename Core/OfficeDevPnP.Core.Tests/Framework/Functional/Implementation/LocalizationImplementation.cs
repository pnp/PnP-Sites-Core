using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class LocalizationImplementation : ImplementationBase
    {
#if !SP2013
        internal void SiteCollectionsLocalization(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                CleanUpTestData(cc);

                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.BaseTemplate = null;
                ptci.PersistMultiLanguageResources = true;
                ptci.FileConnector = new FileSystemConnector(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");
                ptci.HandlersToProcess = Handlers.Fields | Handlers.ContentTypes | Handlers.Lists | Handlers.SupportedUILanguages | Handlers.CustomActions | Handlers.Pages | Handlers.Files | Handlers.Navigation;

                var result = TestProvisioningTemplate(cc, "localization_add.xml", ptci.HandlersToProcess, null, ptci);
                LocalizationValidator validator = new LocalizationValidator(cc.Web);
                Assert.IsTrue(validator.Validate(result.SourceTemplate, result.TargetTemplate, result.SourceTokenParser, result.TargetTokenParser, cc.Web));
            }
        }


        internal void WebLocalization(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                CleanUpTestData(cc);

                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.BaseTemplate = null;
                ptci.PersistMultiLanguageResources = true;
                ptci.FileConnector = new FileSystemConnector(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");
                ptci.HandlersToProcess = Handlers.Fields | Handlers.ContentTypes | Handlers.Lists | Handlers.SupportedUILanguages | Handlers.CustomActions | Handlers.Pages | Handlers.Files;

                var result = TestProvisioningTemplate(cc, "localization_add.xml", ptci.HandlersToProcess, null, ptci);
                LocalizationValidator validator = new LocalizationValidator(cc.Web);
                Assert.IsTrue(validator.Validate(result.SourceTemplate, result.TargetTemplate, result.SourceTokenParser, result.TargetTokenParser, cc.Web));
            }
        }


#region Helper methods
        private void CleanUpTestData(ClientContext cc)
        {
            DeleteLists(cc);
            DeleteContentTypes(cc);
            DeleteCustomActions(cc);
            DeletePages(cc);
        }

        private void DeleteLists(ClientContext cc)
        {
            DeleteListsImplementation(cc);
        }

        private static void DeleteListsImplementation(ClientContext cc)
        {
            cc.Load(cc.Web.Lists, f => f.Include(t => t.DefaultViewUrl));
            cc.ExecuteQueryRetry();

            foreach (var list in cc.Web.Lists.ToList())
            {
                if (list.DefaultViewUrl.Contains("LI_"))
                {
                    list.DeleteObject();
                }
            }
            cc.ExecuteQueryRetry();
        }

        private void DeleteContentTypes(ClientContext cc)
        {
            // Drop the content types
            cc.Load(cc.Web.ContentTypes, f => f.Include(t => t.Group));
            cc.ExecuteQueryRetry();

            foreach (var ct in cc.Web.ContentTypes.ToList())
            {
                if (ct.Group.Equals("PnP Localization Demo"))
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
        private void DeleteCustomActions(ClientContext cc)
        {
            var siteActions = cc.Site.GetCustomActions();
            foreach (var action in siteActions)
            {
                if (action.Name.StartsWith("CA_"))
                {
                    cc.Site.DeleteCustomAction(action.Id);
                }
            }

            var webActions = cc.Web.GetCustomActions();
            foreach (var action in webActions)
            {
                if (action.Name.StartsWith("CA_"))
                {
                    cc.Web.DeleteCustomAction(action.Id);
                }
            }
        }

        private void DeletePages(ClientContext cc)
        {
            Web web = cc.Web;
            web.EnsureProperties(w => w.ServerRelativeUrl);
            string serverRelatedUrl = web.ServerRelativeUrl;

            try
            {
                var file = web.GetFileByServerRelativeUrl(serverRelatedUrl + "/SitePages/LocalizationPage.aspx");
                var file2 = web.GetFileByServerRelativeUrl(serverRelatedUrl + "/SitePages/LocalizationPage2.aspx");
                cc.Load(file);
                cc.Load(file2);
                file.DeleteObject();
                file2.DeleteObject();
                cc.ExecuteQueryRetry();
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName != "System.IO.FileNotFoundException")
                {
                    throw;
                }
            }
        }
#endregion

#endif

    }
}