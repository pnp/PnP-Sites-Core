using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using WebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectPagesTests
    {
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
        private string TestFilePath = "..\\..\\Resources\\office365.png";

        private void DeleteFile(ClientContext ctx, string serverRelativeFileUrl)
        {
            var file = ctx.Web.GetFileByServerRelativeUrl(serverRelativeFileUrl);
            ctx.Load(file, f => f.Exists);
            ctx.ExecuteQueryRetry();

            if (file.Exists)
            {
                file.DeleteObject();
                ctx.ExecuteQueryRetry();
            }

        }


        [TestInitialize()]
        public void Initialize()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Web.EnsureProperty(p => p.ServerRelativeUrl);

                var assetsLibrary = ctx.Web.GetList($"{ctx.Web.ServerRelativeUrl}/SiteAssets");
                ctx.Load(assetsLibrary, p => p.RootFolder);
                ctx.ExecuteQueryRetry();
                var folder = assetsLibrary.RootFolder;

                var fci = new FileCreationInformation();
                fci.Content = System.IO.File.ReadAllBytes(TestFilePath);
                fci.Url = folder.ServerRelativeUrl + "/office365.png";
                fci.Overwrite = true;

                Microsoft.SharePoint.Client.File file = folder.Files.Add(fci);
                ctx.Load(file);
                ctx.ExecuteQueryRetry();
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();

                DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/pagetest.aspx"));
                DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SiteAssets/office365.png"));
            }
        }
        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();
            
            Page page = new Page();
            page.Layout = WikiPageLayout.TwoColumns;
            page.Overwrite = true;
            page.Url = "{site}/sitepages/pagetest.aspx";

           
            var webPart = new WebPart();
            webPart.Column = 1;
            webPart.Row = 1;
            webPart.Contents = webpartcontents;
            webPart.Title = "Script Test";

            page.WebParts.Add(webPart);


            template.Pages.Add(page);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectPages().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();

                var file = ctx.Web.GetFileByServerRelativeUrl(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/pagetest.aspx"));
                ctx.Load(file, f => f.Exists);
                ctx.ExecuteQueryRetry();

                Assert.IsTrue(file.Exists);
                var wps = ctx.Web.GetWebParts(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/pagetest.aspx"));
                Assert.IsTrue(wps.Any());
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
                template = new ObjectPages().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsInstanceOfType(template.Pages, typeof(PageCollection));
            }
        }

        /// <summary>
        /// For some reason this is only an issue in SharePoint 2013 (possibly also in SharePoint 2016).
        /// It appears also to only be an issue for pages/list items and not for sites and lists.
        /// </summary>
        [TestMethod]
        public void CanRemoveCurrentUserAfterBreakingRoleInheritanceWithoutCopyRoleAssignment()
        {
            var template = new ProvisioningTemplate();

            Page page = new Page();
            page.Layout = WikiPageLayout.TwoColumns;
            page.Overwrite = true;
            page.Url = "{site}/sitepages/pagetest.aspx";

            var security = new ObjectSecurity()
            {
                ClearSubscopes = false,
                CopyRoleAssignments = false,
            };

            page.SetSecurity(security);

            template.Pages.Add(page);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var currentUser = ctx.Web.EnsureProperty(w => w.CurrentUser);

                var roleAssignment = new Core.Framework.Provisioning.Model.RoleAssignment();
                roleAssignment.Principal = currentUser.LoginName;
                roleAssignment.Remove = true;

                roleAssignment.RoleDefinition = "{roledefinition:Administrator}";               

                page.Security.RoleAssignments.Add(roleAssignment);

                var parser = new TokenParser(ctx.Web, template);
                new ObjectPages().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();

                var file = ctx.Web.GetFileByServerRelativeUrl(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/pagetest.aspx"));
                ctx.Load(file, f => f.Exists);
                ctx.ExecuteQueryRetry();
                Assert.IsTrue(file.Exists);

                var listItem = file.ListItemAllFields;
                var roleAssignments = listItem.RoleAssignments;
                ctx.Load(roleAssignments);
                ctx.ExecuteQueryRetry();

                Assert.AreEqual(0, roleAssignments.Count);
            }
        }

#if !SP2013 && !SP2016
        [TestMethod]
        public void CanSaveAndLoadHeaderProperties()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();
                var imgUrl = UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SiteAssets/office365.png");

                var pageName = $"{Guid.NewGuid().ToString()}.aspx";
                var newPage = ctx.Web.AddClientSidePage();
                newPage.LayoutType = Pages.ClientSidePageLayoutType.Article;
#if !SP2019
                newPage.PageHeader.TopicHeader = "HEY HEADER";
                newPage.PageHeader.LayoutType = Pages.ClientSidePageHeaderLayoutType.NoImage;
                newPage.PageHeader.ShowTopicHeader = true;
#endif
                newPage.PageHeader.ImageServerRelativeUrl = imgUrl;
                newPage.PageHeader.TranslateX = 1.0;
                newPage.PageHeader.TranslateY = 2.0;
                newPage.Save(pageName);
                newPage.Publish();
                try
                {
                    var readPage = ctx.Web.LoadClientSidePage(pageName);
                    Assert.AreEqual(readPage.LayoutType, Pages.ClientSidePageLayoutType.Article);
#if !SP2019
                    Assert.AreEqual("HEY HEADER", readPage.PageHeader.TopicHeader);
                    Assert.IsTrue(readPage.PageHeader.ShowTopicHeader);
                    Assert.AreEqual(Pages.ClientSidePageHeaderLayoutType.NoImage, readPage.PageHeader.LayoutType);
#endif
                    Assert.AreEqual(imgUrl, readPage.PageHeader.ImageServerRelativeUrl);
                    Assert.AreEqual(1.0, readPage.PageHeader.TranslateX);
                    Assert.AreEqual(2.0, readPage.PageHeader.TranslateY);
                }
                finally
                {
                    DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/" + pageName));
                }
            }
        }
#endif
    }
}
