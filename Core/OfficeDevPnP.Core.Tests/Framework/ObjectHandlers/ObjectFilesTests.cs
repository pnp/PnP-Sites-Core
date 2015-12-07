using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using File = OfficeDevPnP.Core.Framework.Provisioning.Model.File;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectFilesTests
    {
        private string resourceFolder;
        private const string fileName = "ProvisioningTemplate-2015-03-Sample-01.xml";
        private string folder;

        [TestInitialize]
        public void Initialize()
        {
            resourceFolder = string.Format(@"{0}\..\..\Resources\Templates",
                AppDomain.CurrentDomain.BaseDirectory);

            folder = string.Format("test{0}", DateTime.Now.Ticks);
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Web.EnsureProperties(w => w.ServerRelativeUrl);

                var file = ctx.Web.GetFileByServerRelativeUrl(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "test/" + fileName));
                ctx.Load(file, f => f.Exists);
                ctx.ExecuteQueryRetry();
                if (file.Exists)
                {
                    file.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
                if (ctx.Web.RootFolder.FolderExists(folder))
                {
                    var serverFolder = ctx.Web.GetFolderByServerRelativeUrl(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, folder));
                    serverFolder.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();

            FileSystemConnector connector = new FileSystemConnector(resourceFolder, "");

            template.Connector = connector;

            template.Files.Add(new Core.Framework.Provisioning.Model.File() { Overwrite = true, Src = fileName, Folder = folder });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectFiles().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());


                ctx.Web.EnsureProperties(w => w.ServerRelativeUrl);

                var file = ctx.Web.GetFileByServerRelativeUrl(
                    UrlUtility.Combine(ctx.Web.ServerRelativeUrl,
                        UrlUtility.Combine(folder, fileName)));
                ctx.Load(file, f => f.Exists);
                ctx.ExecuteQueryRetry();
                Assert.IsTrue(file.Exists);
            }
        }


        [TestMethod]
        public void CanProvisionObjectsRequiredField()
        {

            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(resourceFolder, "");
            var template = provider.GetTemplate(resourceFolder + "/" + fileName);
            FileSystemConnector connector = new FileSystemConnector(resourceFolder, "");

            template.Connector = connector;
            // replace whatever files is in the template with a file we control
            template.Files.Clear();
            template.Files.Add(new Core.Framework.Provisioning.Model.File() { Overwrite = true, Src = fileName, Folder = "Lists/ProjectDocuments" });

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectField().ProvisionObjects(ctx.Web, template, parser,
                    new ProvisioningTemplateApplyingInformation());
                new ObjectContentType().ProvisionObjects(ctx.Web, template, parser,
                    new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance().ProvisionObjects(ctx.Web, template, parser,
                    new ProvisioningTemplateApplyingInformation());

                new ObjectFiles().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());


                ctx.Web.EnsureProperties(w => w.ServerRelativeUrl);

                var file = ctx.Web.GetFileByServerRelativeUrl(
                    UrlUtility.Combine(ctx.Web.ServerRelativeUrl,
                        UrlUtility.Combine("Lists/ProjectDocuments", fileName)));
                ctx.Load(file, f => f.Exists);
                ctx.ExecuteQueryRetry();
                Assert.IsTrue(file.Exists);

                // cleanup for artifacts specific to this test
                foreach (var list in template.Lists)
                {
                    ctx.Web.GetListByUrl(list.Url).DeleteObject();
                    
                }
            
                foreach (var ct in template.ContentTypes)
                {
                    ctx.Web.GetContentTypeById(ct.Id).DeleteObject();
                }
           
                var idsToDelete = new List<Guid>();
                foreach (var field in ctx.Web.Fields)
                {
                    if (field.Group == "My Columns")
                    {
                        idsToDelete.Add(field.Id);
                    }
                }
                foreach (var guid in idsToDelete)
                {
                    ctx.Web.GetFieldById<Microsoft.SharePoint.Client.Field>(guid).DeleteObject();
                }
                ctx.ExecuteQueryRetry();
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
                template = new ObjectFiles().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsInstanceOfType(template.Files, typeof(List<Core.Framework.Provisioning.Model.File>));
            }
        }
    }
}
