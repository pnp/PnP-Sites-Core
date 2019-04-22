using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201903;
using OfficeDevPnP.Core.Utilities;
using File = System.IO.File;
using ProvisioningTemplate = OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate;

namespace OfficeDevPnP.Core.Tests.Framework.Providers
{
	[TestClass]
	public class XMLSerializer201903Tests
	{
		#region Test variables
		private const string TEST_CATEGORY = "Framework Provisioning XML Serialization\\Deserialization 201903";
		private const string TEST_OUT_FILE = "ProvisioningTemplate-2019-03-Sample-01-test.xml";

		#endregion

		#region Test initialize
		[ClassCleanup]
		public static void Clean()
		{
			var testFilePath = Path.GetFullPath(Path.Combine("../../Resources/Templates", TEST_OUT_FILE));
			if (File.Exists(testFilePath))
			{
				File.Delete(testFilePath);
			}
		}
		#endregion

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Deserialize_Apps()
		{
			var provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
						AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var serializer = new XMLPnPSchemaV201903Serializer();
			var template = provider.GetTemplate("ProvisioningSchema-2019-03-FullSample-01.xml", serializer);

			Assert.AreEqual(3, template.ApplicationLifecycleManagement.Apps.Count);
			Assert.AreEqual(AppAction.Install, template.ApplicationLifecycleManagement.Apps.First().Action);
			Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", template.ApplicationLifecycleManagement.Apps.First().AppId);
			Assert.AreEqual(SyncMode.Synchronously, template.ApplicationLifecycleManagement.Apps.First().SyncMode);
			Assert.AreEqual(AppAction.Update, template.ApplicationLifecycleManagement.Apps[1].Action);
			Assert.AreEqual(AppAction.Uninstall, template.ApplicationLifecycleManagement.Apps[2].Action);
		}

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Serialize_Apps()
		{
			XMLTemplateProvider provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
					AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var result = new ProvisioningTemplate();

			result.ApplicationLifecycleManagement.Apps.Add(new App
			{
				AppId = "1d602ad8-2ef2-4d0b-bfdc-ca71a1d591fc",
				Action = AppAction.Install,
				SyncMode = SyncMode.Asynchronously
			});

			result.ApplicationLifecycleManagement.Apps.Add(new App
			{
				AppId = "6f9a1b33-a13a-4313-b106-0effa15624e6",
				Action = AppAction.Uninstall,
				SyncMode = SyncMode.Synchronously
			});

			result.ApplicationLifecycleManagement.Apps.Add(new App
			{
				AppId = "a44cd745-57ea-44bc-a707-95b985e696e5",
				Action = AppAction.Update,
				SyncMode = SyncMode.Synchronously
			});

			var serializer = new XMLPnPSchemaV201903Serializer();
			provider.SaveAs(result, TEST_OUT_FILE, serializer);

			var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
			Assert.IsTrue(File.Exists(path));
			var xml = XDocument.Load(path);
			Provisioning wrappedResult =
				XMLSerializer.Deserialize<Provisioning>(xml);

			var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
			Assert.AreEqual(3, template.ApplicationLifecycleManagement.Apps.Count());
			Assert.AreEqual("1d602ad8-2ef2-4d0b-bfdc-ca71a1d591fc", template.ApplicationLifecycleManagement.Apps[0].AppId);
			Assert.AreEqual(ApplicationLifecycleManagementAppAction.Install, template.ApplicationLifecycleManagement.Apps[0].Action);
			Assert.AreEqual(ApplicationLifecycleManagementAppAction.Uninstall, template.ApplicationLifecycleManagement.Apps[1].Action);
			Assert.AreEqual(ApplicationLifecycleManagementAppAction.Update, template.ApplicationLifecycleManagement.Apps[2].Action);
			Assert.AreEqual(ApplicationLifecycleManagementAppSyncMode.Asynchronously, template.ApplicationLifecycleManagement.Apps[0].SyncMode);
			Assert.AreEqual(ApplicationLifecycleManagementAppSyncMode.Synchronously, template.ApplicationLifecycleManagement.Apps[1].SyncMode);
		}

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Deserialize_SiteHeader()
		{
			var provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
						AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var serializer = new XMLPnPSchemaV201903Serializer();
			var template = provider.GetTemplate("ProvisioningSchema-2019-03-FullSample-01.xml", serializer);

			Assert.AreEqual(SiteHeaderLayout.Standard, template.Header.Layout);
			Assert.AreEqual(SiteHeaderMenuStyle.MegaMenu, template.Header.MenuStyle);
			Assert.AreEqual(SiteHeaderBackgroundEmphasis.Soft, template.Header.BackgroundEmphasis);
		}

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Serialize_SiteHeader()
		{
			XMLTemplateProvider provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
						AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var result = new ProvisioningTemplate
			{
				Header = new SiteHeader
				{
					MenuStyle = SiteHeaderMenuStyle.Cascading,
					Layout = SiteHeaderLayout.Compact,
					BackgroundEmphasis = SiteHeaderBackgroundEmphasis.Strong
				}
			};

			var serializer = new XMLPnPSchemaV201903Serializer();
			provider.SaveAs(result, TEST_OUT_FILE, serializer);

			var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
			Assert.IsTrue(File.Exists(path));
			var xml = XDocument.Load(path);
			var wrappedResult =
				XMLSerializer.Deserialize<Provisioning>(xml);

			var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

			Assert.AreEqual(HeaderLayout.Compact, template.Header.Layout);
			Assert.AreEqual(HeaderMenuStyle.Cascading, template.Header.MenuStyle);
			Assert.AreEqual(HeaderBackgroundEmphasis.Strong, template.Header.BackgroundEmphasis);
		}

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Deserialize_SiteFooter()
		{
			var provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
						AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var serializer = new XMLPnPSchemaV201903Serializer();
			var template = provider.GetTemplate("ProvisioningSchema-2019-03-FullSample-01.xml", serializer);

			Assert.AreEqual(true, template.Footer.Enabled);
			Assert.AreEqual("logo.png", template.Footer.Logo);
			Assert.AreEqual("FooterName", template.Footer.Name);
			Assert.AreEqual(true, template.Footer.RemoveExistingNodes);
			Assert.AreEqual(6, template.Footer.FooterLinks.Count);
			Assert.AreEqual("www.link01.com", template.Footer.FooterLinks[0].Url);
			Assert.AreEqual("Link 01", template.Footer.FooterLinks[0].DisplayName);
			Assert.AreEqual(3, template.Footer.FooterLinks[3].FooterLinks.Count);
			Assert.AreEqual("www.link04-01.com", template.Footer.FooterLinks[3].FooterLinks[0].Url);
			Assert.AreEqual("Child Link 04-01", template.Footer.FooterLinks[3].FooterLinks[0].DisplayName);
			Assert.IsNull(template.Footer.FooterLinks[3].Url);
			Assert.AreEqual(2, template.Footer.FooterLinks[5].FooterLinks[0].FooterLinks.Count);
		}

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Serialize_SiteFooter()
		{
			XMLTemplateProvider provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
						AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var result = new ProvisioningTemplate
			{
				Footer = new SiteFooter
				{
					Enabled = true,
					Logo = "logo.png",
					Name = "MyFooter",
					RemoveExistingNodes = true,
					FooterLinks = {
						new SiteFooterLink
						{
							Url = "www.link01.com",
							DisplayName = "Link 01"
						},
						new SiteFooterLink
						{
							DisplayName = "Link 02",
							FooterLinks =
							{
								new SiteFooterLink
								{
									Url = "www.link02-01.com",
									DisplayName = "Child Link 01",
								},
								new SiteFooterLink
								{
									Url = "www.link02-02.com",
									DisplayName = "Child Link 02",
								}
							}
						}
					}
				}
			};

			var serializer = new XMLPnPSchemaV201903Serializer();
			provider.SaveAs(result, TEST_OUT_FILE, serializer);

			var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
			Assert.IsTrue(File.Exists(path));
			var xml = XDocument.Load(path);
			var wrappedResult =
				XMLSerializer.Deserialize<Provisioning>(xml);

			var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

			Assert.AreEqual(2, template.Footer.FooterLinks.Count());
			Assert.AreEqual(true, template.Footer.Enabled);
			Assert.AreEqual("logo.png", template.Footer.Logo);
			Assert.AreEqual("MyFooter", template.Footer.Name);
			Assert.AreEqual(true, template.Footer.RemoveExistingNodes);
			Assert.AreEqual("www.link01.com", template.Footer.FooterLinks[0].Url);
			Assert.AreEqual("Link 01", template.Footer.FooterLinks[0].DisplayName);
			Assert.AreEqual(2, template.Footer.FooterLinks[1].FooterLink1.Count());
			Assert.AreEqual("www.link02-01.com", template.Footer.FooterLinks[1].FooterLink1[0].Url);
			Assert.AreEqual("Child Link 01", template.Footer.FooterLinks[1].FooterLink1[0].DisplayName);
		}

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Deserialize_AppCatalog()
		{
			var provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
						AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var serializer = new XMLPnPSchemaV201903Serializer();
			var template = provider.GetTemplate("ProvisioningSchema-2019-03-FullSample-01.xml", serializer);

			Assert.AreEqual(3, template.ApplicationLifecycleManagement.AppCatalog.Packages.Count);
			Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog.Packages[0].Overwrite);
			Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog.Packages[0].SkipFeatureDeployment);
			Assert.AreEqual(PackageAction.UploadAndPublish, template.ApplicationLifecycleManagement.AppCatalog.Packages[0].Action);
			Assert.AreEqual(PackageAction.Publish, template.ApplicationLifecycleManagement.AppCatalog.Packages[1].Action);
			Assert.AreEqual("solution\\spfx-discuss-now.sppkg", template.ApplicationLifecycleManagement.AppCatalog.Packages[0].Src);
			Assert.AreEqual("9672a07b-c111-4a12-bb5b-8d43c2ddd256", template.ApplicationLifecycleManagement.AppCatalog.Packages[2].PackageId);

			Assert.AreEqual(3, template.Tenant.AppCatalog.Packages.Count);
			Assert.AreEqual(true, template.Tenant.AppCatalog.Packages[0].Overwrite);
			Assert.AreEqual(true, template.Tenant.AppCatalog.Packages[0].SkipFeatureDeployment);
			Assert.AreEqual(PackageAction.UploadAndPublish, template.Tenant.AppCatalog.Packages[0].Action);
			Assert.AreEqual(PackageAction.Publish, template.Tenant.AppCatalog.Packages[1].Action);
			Assert.AreEqual("solution\\spfx-discuss-now.sppkg", template.Tenant.AppCatalog.Packages[0].Src);
			Assert.AreEqual("9672a07b-c111-4a12-bb5b-8d43c2ddd256", template.Tenant.AppCatalog.Packages[2].PackageId);
		}

		[TestMethod]
		[TestCategory(TEST_CATEGORY)]
		public void XMLSerializer_Serialize_AppCatalog()
		{
			XMLTemplateProvider provider =
				new XMLFileSystemTemplateProvider(
					String.Format(@"{0}\..\..\Resources",
						AppDomain.CurrentDomain.BaseDirectory),
					"Templates");

			var result = new ProvisioningTemplate();

			var packages = new PackageCollection(result)
			{
				new Package
				{
					SkipFeatureDeployment = true,
					Src = "testpackage.sppkg",
					Overwrite = true,
					Action = PackageAction.Upload,
					PackageId = "60006518-60b3-46d1-8aa7-60a89ce35f03"
				},
				new Package
				{
					SkipFeatureDeployment = true,
					Overwrite = true,
					Action = PackageAction.Publish,
					PackageId = "60006518-60b3-46d1-8aa7-60a89ce35f03"
				}
			};

			result.ApplicationLifecycleManagement = new Core.Framework.Provisioning.Model.ApplicationLifecycleManagement();

			result.ApplicationLifecycleManagement.AppCatalog.Packages.AddRange(packages);
			result.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog, new Core.Framework.Provisioning.Model.ContentDeliveryNetwork());

			var serializer = new XMLPnPSchemaV201903Serializer();
			provider.SaveAs(result, TEST_OUT_FILE, serializer);

			var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
			Assert.IsTrue(File.Exists(path));
			var xml = XDocument.Load(path);
			var wrappedResult =
				XMLSerializer.Deserialize<Provisioning>(xml);

			var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

			Assert.AreEqual(2, template.ApplicationLifecycleManagement.AppCatalog.Count());
			Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog[0].SkipFeatureDeployment);
			Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog[0].Overwrite);
			Assert.AreEqual("testpackage.sppkg", template.ApplicationLifecycleManagement.AppCatalog[0].Src); Assert.AreEqual("60006518-60b3-46d1-8aa7-60a89ce35f03", template.ApplicationLifecycleManagement.AppCatalog[1].PackageId);
			Assert.AreEqual(AppCatalogPackageAction.Publish, template.ApplicationLifecycleManagement.AppCatalog[1].Action);

			Assert.AreEqual(2, wrappedResult.Tenant.AppCatalog.Count());
			Assert.AreEqual(true, wrappedResult.Tenant.AppCatalog[0].SkipFeatureDeployment);
			Assert.AreEqual(true, wrappedResult.Tenant.AppCatalog[0].Overwrite);
			Assert.AreEqual("testpackage.sppkg", wrappedResult.Tenant.AppCatalog[0].Src); Assert.AreEqual("60006518-60b3-46d1-8aa7-60a89ce35f03", wrappedResult.Tenant.AppCatalog[1].PackageId);
			Assert.AreEqual(AppCatalogPackageAction.Publish, wrappedResult.Tenant.AppCatalog[1].Action);
		}
	}
}
