using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectListInstanceTests
    {
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
                var list = ctx.Web.GetListByUrl(string.Format("lists/{0}",listName));
                if (list == null)
                    list = ctx.Web.GetListByUrl(listName);
                if (list != null)
                {
                    list.DeleteObject();
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
            listInstance.TemplateType = (int) ListTemplateType.GenericList;

            Dictionary<string, string> dataValues = new Dictionary<string, string>();
            dataValues.Add("Title","Test");
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
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) {BaseTemplate = ctx.Web.GetBaseTemplate()};

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
    }


}
