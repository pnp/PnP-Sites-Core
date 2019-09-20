using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using File = OfficeDevPnP.Core.Framework.Provisioning.Model.File;
using WebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectFilesTests
    {
        private string resourceFolder;
        private const string fileName = "ProvisioningTemplate-2015-03-Sample-01.xml";
        private string folder;
        private string webpartcontents = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" /><importErrorMessage>Cannot import this Web Part.</importErrorMessage>
    </metaData>
    <data>
      <properties>
        <property name=""ExportMode"" type=""exportmode"">All</property>
        <property name=""HelpUrl"" type=""string"" />
        <property name=""Hidden"" type=""bool"">False</property>
        <property name=""Description"" type=""string"">Allows authors to insert HTML snippets or scripts.</property>
        <property name=""Content"" type=""string"">&lt;script type=""text/javascript""&gt;
alert(""Hello!"");
&lt;/script&gt;</property>
        <property name=""CatalogIconImageUrl"" type=""string"" />
        <property name=""Title"" type=""string"">Script Editor</property>
        <property name=""AllowHide"" type=""bool"">True</property>
        <property name=""AllowMinimize"" type=""bool"">True</property>
        <property name=""AllowZoneChange"" type=""bool"">True</property>
        <property name=""TitleUrl"" type=""string"" />
        <property name=""ChromeType"" type=""chrometype"">None</property>
        <property name=""AllowConnect"" type=""bool"">True</property>
        <property name=""Width"" type=""unit"" />
        <property name=""Height"" type=""unit"" />
        <property name=""HelpMode"" type=""helpmode"">Navigate</property>
        <property name=""AllowEdit"" type=""bool"">True</property>
        <property name=""TitleIconImageUrl"" type=""string"" />
        <property name=""Direction"" type=""direction"">NotSet</property>
        <property name=""AllowClose"" type=""bool"">True</property>
        <property name=""ChromeState"" type=""chromestate"">Normal</property>
      </properties>
    </data>
  </webPart>
</webParts>";

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
                var serverRelativeUrl = UrlUtility.Combine(ctx.Web.ServerRelativeUrl, UrlUtility.Combine(folder, fileName));
                var file = ctx.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
                
                // This call will fail as we're creating a file not bound to a list
                ctx.Load(file);
                try
                {
                    ctx.ExecuteQueryRetry();
                    Assert.IsTrue(file.ServerRelativeUrl.Equals(serverRelativeUrl, StringComparison.InvariantCultureIgnoreCase));
                }
                catch (ServerException ex)
                {
                    // If this throws ServerException (does not belong to list), then shouldn't be trying to set properties)
                    // Handling the exception stating the "The object specified does not belong to a list."
                    if (ex.ServerErrorCode != -2146232832)
                    {
                        throw;
                    }
                }

            }
        }

        [TestMethod]
        public void CanAddWebPartsToForms()
        {
            var template = new ProvisioningTemplate();

            FileSystemConnector connector = new FileSystemConnector(resourceFolder + @"\..", "");

            template.Connector = connector;
            var webPart = new WebPart
            {
                Column = 1,
                Row = 1,
                Contents = webpartcontents,
                Title = "Script Editor",
                Order = 0,
                Zone = "Main"
            };

            var myfile = new Core.Framework.Provisioning.Model.File()
            {
                Overwrite = false,
                Src = "EditForm.aspx",
                Folder = "SitePages/Forms"
            };
            myfile.WebParts.Add(webPart);
            template.Files.Add(myfile);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectFiles().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                ctx.Web.EnsureProperties(w => w.ServerRelativeUrl);

                var file = ctx.Web.GetFileByServerRelativeUrl(
                    UrlUtility.Combine(ctx.Web.ServerRelativeUrl,
                        UrlUtility.Combine("SitePages/Forms", "EditForm.aspx")));
                ctx.Load(file, f => f.Exists);
                ctx.ExecuteQueryRetry();

                // first of all do we even find the form ?
                Assert.IsTrue(file.Exists);
                var webParts = file.GetLimitedWebPartManager(PersonalizationScope.Shared).WebParts;
                ctx.Load(webParts, wp => wp.IncludeWithDefaultProperties(w=>w.Id, w=>w.WebPart, w=>w.WebPart.Title));
                ctx.ExecuteQueryRetry();

                var webPartsArray = webParts.ToArray();
                var webPartExists = false;
                foreach (var webPartDefinition in webPartsArray)
                {
                    if (webPartDefinition.WebPart.Title == "Script Editor")
                    {
                        webPartExists = true;
                        // cleanup after ourselves if we can find the webpart... 
                        webPartDefinition.DeleteWebPart();
                    }
                   
                }
                ctx.ExecuteQueryRetry();
                Assert.IsTrue(webPartExists);
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
                new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser,
                    new ProvisioningTemplateApplyingInformation());
                new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser,
                    new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields).ProvisionObjects(ctx.Web, template, parser,
                    new ProvisioningTemplateApplyingInformation());
                new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings).ProvisionObjects(ctx.Web, template, parser,
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

                Assert.IsInstanceOfType(template.Files, typeof(Core.Framework.Provisioning.Model.FileCollection));
            }
        }
    }
}
