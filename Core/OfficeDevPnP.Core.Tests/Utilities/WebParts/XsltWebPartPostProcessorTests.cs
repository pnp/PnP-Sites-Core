using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Utilities.WebParts.Processors;
using File = Microsoft.SharePoint.Client.File;
using View = Microsoft.SharePoint.Client.View;

namespace OfficeDevPnP.Core.Tests.Utilities.WebParts
{
    [TestClass]
    public class XsltWebPartPostProcessorTests
    {
        private string testPage = "TestDefault.aspx";
        private string folder = "test";

        private static Guid _viewId;
        private static Guid _docsListId;
        private const string ViewName = "Xslt Test View";
        private const string TestListUrl = "Shared Documents";
        private const string ListTitle = "Documents";
        private const string ViewXml = @"
            <Query>
	            <GroupBy Collapse=""TRUE"" GroupLimit=""30"">
		            <FieldRef Name=""Author""/>
	            </GroupBy>
	            <OrderBy>
		            <FieldRef Name=""Created"" Ascending=""FALSE""/>
		            <FieldRef Name=""Title""/>
	            </OrderBy>
	            <Where>
		            <Lt>
			            <FieldRef Name=""Modified""/>
			            <Value Type=""DateTime"">
				            <Today/>
			            </Value>
		            </Lt>
	            </Where>
            </Query>
            <ViewFields>
	            <FieldRef Name=""ID""/>
	            <FieldRef Name=""DocIcon""/>
	            <FieldRef Name=""LinkFilename""/>
	            <FieldRef Name=""Editor""/>
	            <FieldRef Name=""Created""/>
	            <FieldRef Name=""FileSizeDisplay""/>
            </ViewFields>
            <RowLimit Paged=""TRUE"">9</RowLimit>
            <Aggregations Value=""On"">
	            <FieldRef Name=""DocIcon"" Type=""COUNT""/>
            </Aggregations>
            <JSLink>my_link.js</JSLink>
            <Toolbar Type=""Standard""/>";

        [ClassInitialize]
        public static void BeforeAll(TestContext testContext)
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var docsList = ctx.Web.GetListByUrl(TestListUrl);
                docsList.EnsureProperties(d => d.Id);
                var newView = docsList.Views.Add(new ViewCreationInformation
                {
                    Title = ViewName
                });
                newView.ListViewXml = ViewXml;
                newView.Update();
                newView.EnsureProperties(v => v.Id);
                ctx.ExecuteQueryRetry();
                _viewId = newView.Id;
                _docsListId = docsList.Id;
            }
        }

        [ClassCleanup]
        public static void AfterAll()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var docsList = ctx.Web.GetListByUrl(TestListUrl);
                var view = docsList.Views.GetByTitle(ViewName);
                ctx.Load(view);
                view.DeleteObject();
                ctx.ExecuteQueryRetry();
            }
        }

        [TestInitialize]
        public void Initialize()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var resourceFolder = $@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\Resources";
                var template = new ProvisioningTemplate();
                var fileSystemConnector = new FileSystemConnector(resourceFolder, string.Empty);
                template.Connector = fileSystemConnector;
                template.Files.Add(new Core.Framework.Provisioning.Model.File { Overwrite = true, Src = testPage, Folder = folder });
                var parser = new TokenParser(ctx.Web, template);
                new ObjectFiles().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Web.EnsureProperties(w => w.ServerRelativeUrl);

                if (ctx.Web.RootFolder.FolderExists(folder))
                {
                    var serverFolder = ctx.Web.GetFolderByServerRelativeUrl(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, folder));
                    serverFolder.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void UpdatesViewFromViewNameProperty()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""ViewName"" type=""string"">{ViewName}</property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewIdProperty()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""ViewId"" type=""string"">{_viewId}</property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewGuidIdProperty()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""ViewGuid"" type=""string"">{_viewId}</property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewUrlProperty()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""ViewUrl"" type=""string"">{ViewName}.aspx</property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewNameUsingListIdProperty()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListId"" type=""System.Guid, mscorlib, Version = 4.0.0.0, Culture = neutral, PublicKeyToken = b77a5c561934e089"">{_docsListId}</property>
                            <property name=""ViewName"" type=""string"">{ViewName}</property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewNameUsingListNameProperty()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListName"">{_docsListId}</property>
                            <property name=""ViewName"" type=""string"">{ViewName}</property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewNameUsingListDisplayNameProperty()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListDisplayName"" type=""string"">{ListTitle}</property>
                            <property name=""ViewName"" type=""string"">{ViewName}</property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewNameUsingXmlDefinition()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""XmlDefinition"" type=""string""><![CDATA[
                            <View Type=""HTML"" DisplayName=""{ViewName}"" BaseViewID=""1"">
	                            <ViewFields><FieldRef Name=""ID""/></ViewFields>
                            </View>]]></property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewUrlUsingXmlDefinition()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""XmlDefinition"" type=""string""><![CDATA[
                            <View Type=""HTML"" Url=""{ViewName}.aspx"" BaseViewID=""1"">
	                            <ViewFields><FieldRef Name=""ID""/></ViewFields>
                            </View>]]></property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromViewIdUsingXmlDefinition()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""XmlDefinition"" type=""string""><![CDATA[
                            <View Type=""HTML"" Name=""{_viewId}"" BaseViewID=""1"">
	                            <ViewFields><FieldRef Name=""ID""/></ViewFields>
                            </View>]]></property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void UpdatesViewFromXmlDefinition()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""XmlDefinition"" type=""string""><![CDATA[
                            <View Type=""HTML"" Name=""Unknown view name"" BaseViewID=""1"">
	                            {ViewXml}
                            </View>]]></property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            CoreTest(wpXml);
        }

        [TestMethod]
        public void ThrowsInvalidSchemaXmlException()
        {
            var wpXml = $@"
                <webParts>
                      <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                          <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                          <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                          <properties>
                            <property name=""ListUrl"" type=""string"">{TestListUrl}</property>
                            <property name=""XmlDefinition"" type=""string""><![CDATA[
	                            {ViewXml}]]></property>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            try
            {
                CoreTest(wpXml);
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex is ServerException);
            }
        }

        [TestMethod]
        public void UpdatesViewFromChildWebListView()
        {
            Web childWeb = null;
            using (var ctx = TestCommon.CreateClientContext())
            {
                try
                {
                    childWeb = CreateTestTeamSubSite(ctx.Web);
                    childWeb.EnsureProperty(w => w.Id);
                    var docsList = childWeb.GetListByUrl(TestListUrl);
                    docsList.EnsureProperties(l => l.Id);
                    var wpXml =
                    $@"<webParts>
                        <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                        <metaData>
                            <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
                            <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                        </metaData>
                        <data>
                            <properties>
                            <property name=""ListId"" type=""System.Guid, mscorlib, Version = 4.0.0.0, Culture = neutral, PublicKeyToken = b77a5c561934e089"">{docsList.Id}</property>
                            <property name=""WebId"" type=""System.Guid, mscorlib, Version = 4.0.0.0, Culture = neutral, PublicKeyToken = b77a5c561934e089"">{childWeb.Id}</property>
                            <property name=""XmlDefinition"" type=""string""><![CDATA[
	                            <View Type=""HTML"" Name=""Unknown view name"" BaseViewID=""1"">
	                                {ViewXml}
                                </View>]]></property>
                            </properties>
                        </data>
                        </webPart>
                    </webParts>";

                    var webParts = CreateBasicWebPart(wpXml);

                    var file = GetFile(ctx);
                    var wpDefinition = AddWebPart(file, wpXml);

                    var xsltPostProcessor = new XsltWebPartPostProcessor(webParts.WebPart);
                    xsltPostProcessor.Process(wpDefinition, file);

                    var view = childWeb.GetListByUrl(TestListUrl).GetViewById(wpDefinition.Id);

                    AssertViewIsValid(view);
                }
                finally
                {
                    if (childWeb != null && childWeb.Context != null)
                    {
                        childWeb.DeleteObject();
                        childWeb.Context.ExecuteQueryRetry();
                        childWeb.Context.Dispose();
                    }
                }
            }
        }

        private void CoreTest(string wpXml)
        {
            var webParts = CreateBasicWebPart(wpXml);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var file = GetFile(ctx);
                var wpDefinition = AddWebPart(file, wpXml);

                var xsltPostProcessor = new XsltWebPartPostProcessor(webParts.WebPart);
                xsltPostProcessor.Process(wpDefinition, file);

                var view = ctx.Web.GetListByUrl(TestListUrl).GetViewById(wpDefinition.Id);

                AssertViewIsValid(view);
            }
        }

        private void AssertViewIsValid(View view)
        {
            view.EnsureProperties(v => v.ViewFields);

            Assert.AreEqual(view.JSLink, "my_link.js");
            Assert.AreEqual(view.Aggregations, @"<FieldRef Name=""DocIcon"" Type=""COUNT"" />");
            Assert.AreEqual(view.AggregationsStatus, "On");
            Assert.AreEqual(view.RowLimit, (uint)9);
            Assert.IsTrue(view.ViewFields.Count == 6);
            Assert.IsNotNull(view.ViewFields.SingleOrDefault(f => f.Equals("FileSizeDisplay")));
            Assert.IsTrue(view.ViewQuery.IndexOf(@"<GroupBy Collapse=""TRUE"" GroupLimit=""30"">", StringComparison.OrdinalIgnoreCase) != -1);
            Assert.IsTrue(view.ViewQuery.IndexOf(@"<FieldRef Name=""Created"" Ascending=""FALSE""", StringComparison.OrdinalIgnoreCase) != -1);
            Assert.IsTrue(view.Hidden);
        }

        private File GetFile(ClientContext ctx)
        {
            ctx.Web.EnsureProperties(w => w.ServerRelativeUrl);

            var serverRelativeUrl = UrlUtility.Combine(ctx.Web.ServerRelativeUrl, $"{folder}/{testPage}");
            var webPartPage = ctx.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
            ctx.Web.Context.Load(webPartPage);
            ctx.Web.Context.ExecuteQueryRetry();

            return webPartPage;
        }

        private WebPartDefinition AddWebPart(File webPartPage, string wpXml)
        {
            var limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var oWebPartDefinition = limitedWebPartManager.ImportWebPart(wpXml);

            var wpdNew = limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, "wpzMain", 1);
            webPartPage.Context.Load(wpdNew);
            webPartPage.Context.ExecuteQueryRetry();

            return wpdNew;
        }

        private static Stream GetXmlStream(string xml)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(xml);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        private Core.Utilities.WebParts.Schema.WebParts CreateBasicWebPart(string wpXml)
        {
            var serializer = new XmlSerializer(typeof(Core.Utilities.WebParts.Schema.WebParts));
            using (var xmlStream = GetXmlStream(wpXml))
            using (var xmlReader = new XmlTextReader(xmlStream))
            {
                xmlReader.Namespaces = false;
                return (Core.Utilities.WebParts.Schema.WebParts)serializer.Deserialize(xmlReader);
            }
        }

        private Web CreateTestTeamSubSite(Web parentWeb)
        {
            var siteUrl = GetRandomString();
            var webInfo = new WebCreationInformation
            {
                Title = siteUrl,
                Url = siteUrl,
                Description = siteUrl,
                Language = 1033,
                UseSamePermissionsAsParentSite = true,
                WebTemplate = "STS#0"
            };

            var web = parentWeb.Webs.Add(webInfo);
            parentWeb.Context.Load(web);
            parentWeb.Context.ExecuteQueryRetry();

            var ctxTestTeamSubSite = parentWeb.Context.Clone(TestCommon.DevSiteUrl + "/" + siteUrl);
            return ctxTestTeamSubSite.Web;
        }

        private string GetRandomString()
        {
            var chars = "abcdefghijklmnopqrstuvwxyz";
            var random = new Random();
            return new string(Enumerable.Repeat(chars, 4).Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
