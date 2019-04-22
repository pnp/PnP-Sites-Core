using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.AzureActiveDirectory;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201903;
using OfficeDevPnP.Core.Utilities;
using File = System.IO.File;
using ProvisioningTemplate = OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate;
using TeamTemplate = OfficeDevPnP.Core.Framework.Provisioning.Model.Teams.TeamTemplate;

namespace OfficeDevPnP.Core.Tests.Framework.Providers
{
    /// <summary>
    /// Covers below objects:
    /// ProvisioningTemplate:
    ///     ALM
    ///     Header
    ///     Footer
    ///     ProvisioningTemplateWebhooks 
    /// Teams:
    ///     TeamTemplate
    ///     Team
    ///     Apps
    /// AzureActiveDirectory
    ///     Users
    /// Tenant
    ///     AppCatalog
    ///     WebApiPermissions
    ///     ContentDeliveryNetwork
    ///     SiteDesigns
    ///     SiteScripts
    ///     StorageEntities
    ///     Themes
    /// </summary>
    [TestClass]
    public class XMLSerializer201903Tests
    {
        #region Test variables
        private const string TEST_CATEGORY = "Framework Provisioning XML Serialization\\Deserialization 201903";
        private const string TEST_OUT_FILE = "ProvisioningTemplate-2019-03-Sample-01-test.xml";
        private const string TEST_TEMPLATE = "ProvisioningSchema-2019-03-FullSample-01.xml";

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
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var serializer = new XMLPnPSchemaV201903Serializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

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
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

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
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var serializer = new XMLPnPSchemaV201903Serializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(SiteHeaderLayout.Standard, template.Header.Layout);
            Assert.AreEqual(SiteHeaderMenuStyle.MegaMenu, template.Header.MenuStyle);
            Assert.AreEqual(SiteHeaderBackgroundEmphasis.Soft, template.Header.BackgroundEmphasis);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SiteHeader()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

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
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var serializer = new XMLPnPSchemaV201903Serializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

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
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

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
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var serializer = new XMLPnPSchemaV201903Serializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

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
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

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
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);

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

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ProvisioningTemplateWebhook()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var serializer = new XMLPnPSchemaV201903Serializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(2, template.ProvisioningTemplateWebhooks.Count);
            Assert.IsTrue(template.ProvisioningTemplateWebhooks[0].Url.Contains("https://mywebhook.azurefunctions.net/function01"));
            Assert.AreEqual(ProvisioningTemplateWebhookMethod.GET, template.ProvisioningTemplateWebhooks[0].Method);
            Assert.AreEqual(ProvisioningTemplateWebhookKind.ProvisioningStarted, template.ProvisioningTemplateWebhooks[0].Kind);
            Assert.AreEqual(ProvisioningTemplateWebhookBodyFormat.Json, template.ProvisioningTemplateWebhooks[1].BodyFormat);
            Assert.AreEqual(true, template.ProvisioningTemplateWebhooks[1].Async);
            Assert.AreEqual(3, template.ProvisioningTemplateWebhooks[0].Parameters.Count);
            Assert.IsTrue(template.ProvisioningTemplateWebhooks[0].Parameters.ContainsKey("Param01"));
            Assert.AreEqual("Value01", template.ProvisioningTemplateWebhooks[0].Parameters["Param01"]);
            Assert.AreEqual("{sitecollection}", template.ProvisioningTemplateWebhooks[1].Parameters["Site"]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ProvisioningTemplateWebhook()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var result = new ProvisioningTemplate();

            result.ProvisioningTemplateWebhooks.Add(new ProvisioningTemplateWebhook
            {
                Url = "https://my.url/func01",
                Method = ProvisioningTemplateWebhookMethod.GET,
                Async = false,
                Kind = ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted,
                BodyFormat = ProvisioningTemplateWebhookBodyFormat.Xml,
                Parameters = new Dictionary<string, string>
                {
                    {"Param01", "Value01"},
                    {"Param02", "Value01"},
                }
            });

            result.ProvisioningTemplateWebhooks.Add(new ProvisioningTemplateWebhook
            {
                Url = "https://my.url/func01",
                Method = ProvisioningTemplateWebhookMethod.POST,
                Async = true,
                Kind = ProvisioningTemplateWebhookKind.ProvisioningCompleted,
                BodyFormat = ProvisioningTemplateWebhookBodyFormat.FormUrlEncoded,
                Parameters = new Dictionary<string, string>
                {
                    {"Param01", "Value01"},
                    {"Param02", "Value01"},
                }
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.AreEqual(2, template.ProvisioningTemplateWebhooks.Count());
            Assert.AreEqual("https://my.url/func01", template.ProvisioningTemplateWebhooks[0].Url);
            Assert.AreEqual(false, template.ProvisioningTemplateWebhooks[0].Async);
            Assert.AreEqual(ProvisioningTemplateWebhooksProvisioningTemplateWebhookBodyFormat.Xml, template.ProvisioningTemplateWebhooks[0].BodyFormat);
            Assert.AreEqual(ProvisioningTemplateWebhooksProvisioningTemplateWebhookMethod.GET, template.ProvisioningTemplateWebhooks[0].Method);
            Assert.AreEqual(ProvisioningTemplateWebhooksProvisioningTemplateWebhookMethod.POST, template.ProvisioningTemplateWebhooks[1].Method);
            Assert.AreEqual(ProvisioningTemplateWebhooksProvisioningTemplateWebhookKind.ProvisioningCompleted, template.ProvisioningTemplateWebhooks[1].Kind);
            Assert.AreEqual(ProvisioningTemplateWebhooksProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted, template.ProvisioningTemplateWebhooks[0].Kind);
            Assert.AreEqual(2, template.ProvisioningTemplateWebhooks[0].Parameters.Count());
            Assert.AreEqual("Param01", template.ProvisioningTemplateWebhooks[0].Parameters[0].Key);
            Assert.AreEqual("Value01", template.ProvisioningTemplateWebhooks[0].Parameters[0].Value);
            Assert.AreEqual("Param01", template.ProvisioningTemplateWebhooks[1].Parameters[0].Key);
            Assert.AreEqual("Value01", template.ProvisioningTemplateWebhooks[1].Parameters[0].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TeamTemplate()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var teamsTemplate = hierarchy.Teams.TeamTemplates;

            Assert.AreEqual(1, teamsTemplate.Count);
            Assert.AreEqual("Custom", teamsTemplate.First().Classification);
            Assert.AreEqual(TeamVisibility.Public, teamsTemplate.First().Visibility);
            Assert.IsTrue(teamsTemplate.First().JsonTemplate.Contains("here goes the team template JSON"));
            Assert.AreEqual("Sample Team from Template", teamsTemplate.First().Description);
            Assert.AreEqual("Team from Template 01", teamsTemplate.First().DisplayName);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TeamTemplate()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Teams.TeamTemplates.Add(new TeamTemplate
            {
                Classification = "MyClass",
                JsonTemplate = "{JSON template here}",
                Description = "Desc",
                Visibility = TeamVisibility.Private,
                DisplayName = "MyTemplate"
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);

            var teamTempaltes = wrappedResult.Teams.Items
                .Where(t => t is Core.Framework.Provisioning.Providers.Xml.V201903.TeamTemplate).Cast<Core.Framework.Provisioning.Providers.Xml.V201903.TeamTemplate>().ToList();

            Assert.AreEqual(1, teamTempaltes.Count);
            Assert.AreEqual("MyClass", teamTempaltes[0].Classification);
            Assert.AreEqual("Desc", teamTempaltes[0].Description);
            Assert.AreEqual("MyTemplate", teamTempaltes[0].DisplayName);
            Assert.AreEqual(BaseTeamVisibility.Private, teamTempaltes[0].Visibility);
            Assert.AreEqual(true, teamTempaltes[0].VisibilitySpecified);
            Assert.IsTrue(teamTempaltes[0].Text[0].Contains("JSON template here"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Team()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var teams = hierarchy.Teams.Teams;

            Assert.AreEqual(2, teams.Count);

            // team common properties
            Assert.AreEqual("Sample Team 01", teams[0].DisplayName);
            Assert.AreEqual("This is just a sample Team 01", teams[0].Description);
            Assert.AreEqual("{TeamId:GroupMailNickname}", teams[0].CloneFrom);
            Assert.AreEqual("{groupid:DisplayName}", teams[1].GroupId);
            Assert.AreEqual("Private", teams[1].Classification);
            Assert.AreEqual(TeamSpecialization.EducationStandard, teams[1].Specialization);
            Assert.AreEqual(TeamVisibility.Public, teams[1].Visibility);
            Assert.AreEqual(false, teams[1].Archived);
            Assert.AreEqual("sample.group", teams[1].MailNickname);

            // team security
            var security = teams[0].Security;
            Assert.AreEqual(false, security.ClearExistingMembers);
            Assert.AreEqual(true, security.ClearExistingOwners);
            Assert.AreEqual(2, security.Owners.Count);
            Assert.AreEqual("owner01@domain.onmicrosoft.com", security.Owners[0].UserPrincipalName);
            Assert.AreEqual(3, security.Members.Count);
            Assert.AreEqual("user01@domain.onmicrosoft.com", security.Members[0].UserPrincipalName);

            // team fun settings
            Assert.AreEqual(true, teams[1].FunSettings.AllowCustomMemes);
            Assert.AreEqual(true, teams[1].FunSettings.AllowGiphy);
            Assert.AreEqual(true, teams[1].FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(TeamGiphyContentRating.Strict, teams[1].FunSettings.GiphyContentRating);

            // team guest settings
            Assert.AreEqual(true, teams[1].GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(false, teams[1].GuestSettings.AllowDeleteChannels);

            // team memebers settings
            Assert.AreEqual(false, teams[1].MemberSettings.AllowDeleteChannels);
            Assert.AreEqual(true, teams[1].MemberSettings.AllowAddRemoveApps);
            Assert.AreEqual(true, teams[1].MemberSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, teams[1].MemberSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(false, teams[1].MemberSettings.AllowCreateUpdateRemoveTabs);

            // team messaging settings
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(false, teams[1].MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowUserEditMessages);

            // team channels
            var channels = teams[1].Channels;
            Assert.AreEqual(3, channels.Count);
            Assert.AreEqual("This is just a Sample Channel", channels[0].Description);
            Assert.AreEqual("Sample Channel 01", channels[0].DisplayName);
            Assert.AreEqual(true, channels[0].IsFavoriteByDefault);
            Assert.AreEqual(1, channels[0].Tabs.Count);
            Assert.AreEqual("My Tab 01", channels[0].Tabs[0].DisplayName);
            Assert.AreEqual("12345", channels[0].Tabs[0].TeamsAppId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView", channels[0].Tabs[0].Configuration.ContentUrl);
            Assert.AreEqual("2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[0].Configuration.EntityId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/uninstallTab", channels[0].Tabs[0].Configuration.RemoveUrl);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[0].Configuration.WebsiteUrl);
            Assert.IsTrue(channels[0].TabResources[0].TabResourceSettings.Contains("\"displayName\": \"Notebook name\""));
            Assert.AreEqual("{TeamsTabId:TabDisplayName}", channels[0].TabResources[0].TargetTabId);
            Assert.AreEqual(TabResourceType.Notebook, channels[0].TabResources[0].Type);
            Assert.AreEqual(1, channels[0].Messages.Count);
            Assert.IsTrue(channels[0].Messages[0].Message.Contains("Welcome to this channel"));

            // team apps
            Assert.AreEqual(2, teams[1].Apps.Count);
            Assert.AreEqual("12345678-9abc-def0-123456789a", teams[1].Apps[0].AppId);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Team()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Teams.Teams.Add(new Team
            {
                DisplayName = "Sample Team 01",
                Description = "This is just a sample Team 01",
                CloneFrom = "{TeamId:GroupMailNickname}",
                Archived = false,
                MailNickname = "sample.group",
                GroupId = "12345",
                Classification = "Private",
                Specialization = TeamSpecialization.EducationClass,
                Visibility = TeamVisibility.Private,
                FunSettings = new TeamFunSettings
                {
                    AllowCustomMemes = true,
                    AllowGiphy = true,
                    GiphyContentRating = TeamGiphyContentRating.Moderate,
                    AllowStickersAndMemes = true
                },
                GuestSettings = new TeamGuestSettings
                {
                    AllowDeleteChannels = true,
                    AllowCreateUpdateChannels = true
                },
                MemberSettings = new TeamMemberSettings
                {
                    AllowDeleteChannels = true,
                    AllowCreateUpdateChannels = true,
                    AllowCreateUpdateRemoveTabs = true,
                    AllowCreateUpdateRemoveConnectors = true,
                    AllowAddRemoveApps = true
                },
                MessagingSettings = new TeamMessagingSettings
                {
                    AllowChannelMentions = true,
                    AllowTeamMentions = true,
                    AllowUserEditMessages = true,
                    AllowOwnerDeleteMessages = true,
                    AllowUserDeleteMessages = true
                },
                Security = new Core.Framework.Provisioning.Model.Teams.TeamSecurity
                {
                    ClearExistingMembers = true,
                    ClearExistingOwners = true,
                    Members = {
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "user01@domain.onmicrosoft.com"
                        },
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "user02@domain.onmicrosoft.com"
                        }
                    },
                    Owners =
                    {
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "owner01@domain.onmicrosoft.com"
                        },
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "owner02@domain.onmicrosoft.com"
                        }
                    }
                },
                Apps =
                {
                    new TeamAppInstance
                    {
                        AppId = "12345678-9abc-def0-123456789a"
                    }
                },
                Channels =
                {
                    new Core.Framework.Provisioning.Model.Teams.TeamChannel
                    {
                        Description = "This is just a Sample Channel",
                        DisplayName = "Sample Channel 01",
                        IsFavoriteByDefault = true,
                        Tabs =
                        {
                            new TeamTab
                            {
                                DisplayName = "My Tab 01",
                                TeamsAppId = "12345",
                                Configuration = new TeamTabConfiguration
                                {
                                    ContentUrl = "https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView",
                                    WebsiteUrl = "https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154",
                                    EntityId = "2DCA2E6C7A10415CAF6B8AB6661B3154",
                                    RemoveUrl = "https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/uninstallTab"
                                }
                            }
                        },
                        TabResources =
                        {
                            new TeamTabResource
                            {
                                TabResourceSettings = "{\"displayName\": \"Notebook name\"}",
                                TargetTabId = "{TeamsTabId:TabDisplayName}",
                                Type = TabResourceType.Planner
                            }
                        },
                        Messages =
                        {
                            new TeamChannelMessage
                            {
                                Message = "{\"body\": \"message\"}"
                            }
                        }
                    }
                }
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);

            var teams = wrappedResult.Teams.Items
                .Where(t => t is TeamWithSettings).Cast<TeamWithSettings>().ToList();

            Assert.AreEqual(1, teams.Count);
            var team = teams[0];

            // team common properties
            Assert.AreEqual("Sample Team 01", team.DisplayName);
            Assert.AreEqual("This is just a sample Team 01", team.Description);
            Assert.AreEqual("{TeamId:GroupMailNickname}", team.CloneFrom);
            Assert.AreEqual("12345", team.GroupId);
            Assert.AreEqual("Private", team.Classification);
            Assert.AreEqual(TeamWithSettingsSpecialization.EducationClass, team.Specialization);
            Assert.AreEqual(BaseTeamVisibility.Private, team.Visibility);
            Assert.AreEqual(false, team.Archived);
            Assert.AreEqual("sample.group", team.MailNickname);

            // team security
            var security = team.Security;
            Assert.AreEqual(true, security.Members.ClearExistingItems);
            Assert.AreEqual(true, security.Owners.ClearExistingItems);
            Assert.AreEqual(2, security.Owners.User.Count());
            Assert.AreEqual("owner01@domain.onmicrosoft.com", security.Owners.User[0].UserPrincipalName);
            Assert.AreEqual(2, security.Members.User.Count());
            Assert.AreEqual("user01@domain.onmicrosoft.com", security.Members.User[0].UserPrincipalName);

            // team fun settings
            Assert.AreEqual(true, team.FunSettings.AllowCustomMemes);
            Assert.AreEqual(true, team.FunSettings.AllowGiphy);
            Assert.AreEqual(true, team.FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(TeamWithSettingsFunSettingsGiphyContentRating.Moderate, team.FunSettings.GiphyContentRating);

            // team guest settings
            Assert.AreEqual(true, team.GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.GuestSettings.AllowDeleteChannels);

            // team memebers settings
            Assert.AreEqual(true, team.MembersSettings.AllowDeleteChannels);
            Assert.AreEqual(true, team.MembersSettings.AllowAddRemoveApps);
            Assert.AreEqual(true, team.MembersSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.MembersSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(true, team.MembersSettings.AllowCreateUpdateRemoveTabs);

            // team messaging settings
            Assert.AreEqual(true, team.MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(true, team.MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserEditMessages);

            // team channels
            var channels = team.Channels;
            Assert.AreEqual(1, channels.Count());
            Assert.AreEqual("This is just a Sample Channel", channels[0].Description);
            Assert.AreEqual("Sample Channel 01", channels[0].DisplayName);
            Assert.AreEqual(true, channels[0].IsFavoriteByDefault);
            Assert.AreEqual(1, channels[0].Tabs.Count());
            Assert.AreEqual("My Tab 01", channels[0].Tabs[0].DisplayName);
            Assert.AreEqual("12345", channels[0].Tabs[0].TeamsAppId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView", channels[0].Tabs[0].Configuration.ContentUrl);
            Assert.AreEqual("2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[0].Configuration.EntityId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/uninstallTab", channels[0].Tabs[0].Configuration.RemoveUrl);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[0].Configuration.WebsiteUrl);
            Assert.IsTrue(channels[0].TabResources[0].TabResourceSettings.Contains("\"displayName\": \"Notebook name\""));
            Assert.AreEqual("{TeamsTabId:TabDisplayName}", channels[0].TabResources[0].TargetTabId);
            Assert.AreEqual(TeamTabResourcesTabResourceType.Planner, channels[0].TabResources[0].Type);

            // team apps
            Assert.AreEqual(1, team.Apps.Count());
            Assert.AreEqual("12345678-9abc-def0-123456789a", team.Apps[0].AppId);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TeamApps()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var apps = hierarchy.Teams.Apps;

            Assert.AreEqual(2, apps.Count);
            Assert.AreEqual("APP001", apps[0].AppId);
            Assert.AreEqual("./custom-app-01.json", apps[0].PackageUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TeamApps()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Teams.Apps.Add(new TeamApp
            {
                AppId = "APP001",
                PackageUrl = "./custom-app-02.zip"
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);

            var teamApps = wrappedResult.Teams.Apps;

            Assert.AreEqual(1, teamApps.Count());
            Assert.AreEqual("APP001", teamApps[0].AppId);
            Assert.AreEqual("./custom-app-02.zip", teamApps[0].PackageUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_AzureAD()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var users = hierarchy.AzureActiveDirectory.Users;

            Assert.AreEqual(2, users.Count);
            Assert.AreEqual("John White", users[0].DisplayName);
            Assert.AreEqual(true, users[0].AccountEnabled);
            Assert.AreEqual("john.white", users[0].MailNickname);
            Assert.AreEqual("john.white@{parameter:domain}.onmicrosoft.com", users[0].UserPrincipalName);
            Assert.AreEqual("Policy1", users[0].PasswordPolicies);
            Assert.AreEqual("photo.jpg", users[0].ProfilePhoto);

            var passWord = new SecureString();

            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            Assert.AreEqual(true, users[0].PasswordProfile.ForceChangePasswordNextSignIn);
            Assert.AreEqual(false, users[0].PasswordProfile.ForceChangePasswordNextSignInWithMfa);
            Assert.IsFalse(users[0].PasswordProfile.Password == null);
            Assert.AreEqual(2, users[0].Licenses.Count);
            Assert.AreEqual("26d45bd9-adf1-46cd-a9e1-51e9a5524128", users[0].Licenses[0].SkuId);
            Assert.AreEqual("e212cbc7-0961-4c40-9825-01117710dcb1", users[0].Licenses[0].DisabledPlans[0]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_AzureAD()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };
            result.ParentHierarchy.AzureActiveDirectory.Users.Add(new Core.Framework.Provisioning.Model.AzureActiveDirectory.User
            {
                AccountEnabled = true,
                DisplayName = "John White",
                MailNickname = "john.white",
                UserPrincipalName = "john.white@{parameter:domain}.onmicrosoft.com",
                PasswordPolicies = "Policy1",
                ProfilePhoto = "photo.jpg",
                PasswordProfile = new PasswordProfile
                {
                    ForceChangePasswordNextSignIn = true,
                    ForceChangePasswordNextSignInWithMfa = true,
                    Password = passWord
                },
                Licenses =
                {
                    new UserLicense
                    {
                        SkuId = "26d45bd9-adf1-46cd-a9e1-51e9a5524128",
                        DisabledPlans = new []{ "e212cbc7-0961-4c40-9825-01117710dcb1", "3fb82609-8c27-4f7b-bd51-30634711ee67", "b1188c4c-1b36-4018-b48b-ee07604f6feb" }
                    },
                    new UserLicense
                    {
                        SkuId = "26d45bd9-adf1-46cd-a9e1-51e9a5524128"
                    }
                }
            });


            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var users = wrappedResult.AzureActiveDirectory.Users;

            Assert.AreEqual(1, users.Count());
            Assert.AreEqual("John White", users[0].DisplayName);
            Assert.AreEqual(true, users[0].AccountEnabled);
            Assert.AreEqual("john.white", users[0].MailNickname);
            Assert.AreEqual("john.white@{parameter:domain}.onmicrosoft.com", users[0].UserPrincipalName);
            Assert.AreEqual("Policy1", users[0].PasswordPolicies);
            Assert.AreEqual("photo.jpg", users[0].ProfilePhoto);

            Assert.AreEqual(true, users[0].PasswordProfile.ForceChangePasswordNextSignIn);
            Assert.AreEqual(true, users[0].PasswordProfile.ForceChangePasswordNextSignInWithMfa);
            Assert.IsFalse(users[0].PasswordProfile.Password == null);
            Assert.AreEqual(2, users[0].Licenses.Count());
            Assert.AreEqual("26d45bd9-adf1-46cd-a9e1-51e9a5524128", users[0].Licenses[0].SkuId);
            Assert.AreEqual("e212cbc7-0961-4c40-9825-01117710dcb1", users[0].Licenses[0].DisabledPlans[0]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_AppCatalog()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var appCatalog = hierarchy.Tenant.AppCatalog;

            Assert.AreEqual(3, appCatalog.Packages.Count);
            Assert.AreEqual("solution\\spfx-discuss-now.sppkg", appCatalog.Packages[0].Src);
            Assert.AreEqual(PackageAction.UploadAndPublish, appCatalog.Packages[0].Action);
            Assert.AreEqual(true, appCatalog.Packages[0].Overwrite);
            Assert.AreEqual(true, appCatalog.Packages[0].SkipFeatureDeployment);
            Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", appCatalog.Packages[1].PackageId);
            Assert.AreEqual(PackageAction.Publish, appCatalog.Packages[1].Action);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Seserialize_Tenant_AppCatalog()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };
            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog, new Core.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.AppCatalog.Packages.Add(new Package
            {
                Action = PackageAction.Publish,
                Src = "solution\\spfx-discuss-now.sppkg",
                SkipFeatureDeployment = true,
                Overwrite = true
            });
            result.Tenant.AppCatalog.Packages.Add(new Package
            {
                Action = PackageAction.Upload,
                SkipFeatureDeployment = true,
                Overwrite = true,
                PackageId = "d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27"
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var packages = wrappedResult.Tenant.AppCatalog;

            Assert.AreEqual(2, packages.Count());
            Assert.AreEqual("solution\\spfx-discuss-now.sppkg", packages[0].Src);
            Assert.AreEqual(AppCatalogPackageAction.Publish, packages[0].Action);
            Assert.AreEqual(true, packages[0].Overwrite);
            Assert.AreEqual(true, packages[0].SkipFeatureDeployment);
            Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", packages[1].PackageId);
            Assert.AreEqual(AppCatalogPackageAction.Upload, packages[1].Action);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_WebApiPermissions()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var apiPermission = hierarchy.Tenant.WebApiPermissions;

            Assert.AreEqual("Microsoft.Graph", apiPermission[0].Resource);
            Assert.AreEqual("User.ReadBasic.All", apiPermission[0].Scope);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Seserialize_Tenant_WebApiPermissions()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };
            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog, new Core.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.WebApiPermissions.Add(
                new WebApiPermission
                {
                    Resource = "Microsoft.Graph",
                    Scope = "User.ReadBasic.All"
                }
            );

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var apiPermissions = wrappedResult.Tenant.WebApiPermissions;

            Assert.AreEqual("Microsoft.Graph", apiPermissions[0].Resource);
            Assert.AreEqual("User.ReadBasic.All", apiPermissions[0].Scope);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_ContentDeliveryNetwork()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var cdn = hierarchy.Tenant.ContentDeliveryNetwork;

            Assert.AreEqual(true, cdn.PublicCdn.Enabled);
            Assert.AreEqual("JS,CSS", cdn.PublicCdn.ExcludeIfNoScriptDisabled);
            Assert.AreEqual("HBI,GDPR", cdn.PublicCdn.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.PublicCdn.IncludeFileExtensions);
            Assert.AreEqual(true, cdn.PublicCdn.NoDefaultOrigins);
            Assert.AreEqual(OriginAction.Add, cdn.PublicCdn.Origins[0].Action);
            Assert.AreEqual("sites/CDN/CDNFiles", cdn.PublicCdn.Origins[0].Url);

            Assert.AreEqual(true, cdn.PrivateCdn.Enabled);
            Assert.AreEqual("HIB,GDPR", cdn.PrivateCdn.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.PrivateCdn.IncludeFileExtensions);
            Assert.AreEqual(false, cdn.PrivateCdn.NoDefaultOrigins);
            Assert.AreEqual(OriginAction.Add, cdn.PrivateCdn.Origins[0].Action);
            Assert.AreEqual("sites/CDNPrivate/CDNFiles", cdn.PrivateCdn.Origins[0].Url);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Seserialize_Tenant_ContentDeliveryNetwork()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            var cdnSettings = new CdnSettings
            {
                Enabled = true,
                NoDefaultOrigins = true,
                ExcludeIfNoScriptDisabled = "JS,CSS",
                ExcludeRestrictedSiteClassifications = "HBI,GDPR",
                IncludeFileExtensions = "PDF,XML,JPG,JS,PNG",
            };
            cdnSettings.Origins.Add(new CdnOrigin
            {
                Action = OriginAction.Add,
                Url = "sites/CDN/CDNFiles"
            });
            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new Core.Framework.Provisioning.Model.ContentDeliveryNetwork(cdnSettings, cdnSettings));

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var cdn = wrappedResult.Tenant.ContentDeliveryNetwork;

            Assert.AreEqual(true, cdn.Public.Enabled);
            Assert.AreEqual("JS,CSS", cdn.Public.ExcludeIfNoScriptDisabled);
            Assert.AreEqual("HBI,GDPR", cdn.Public.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.Public.IncludeFileExtensions);
            Assert.AreEqual(true, cdn.Public.NoDefaultOrigins);
            Assert.AreEqual(CdnSettingOriginAction.Add, cdn.Public.Origins[0].Action);
            Assert.AreEqual("sites/CDN/CDNFiles", cdn.Public.Origins[0].Url);

            Assert.AreEqual(true, cdn.Private.Enabled);
            Assert.AreEqual("HBI,GDPR", cdn.Private.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.Private.IncludeFileExtensions);
            Assert.AreEqual(true, cdn.Private.NoDefaultOrigins);
            Assert.AreEqual(CdnSettingOriginAction.Add, cdn.Private.Origins[0].Action);
            Assert.AreEqual("sites/CDN/CDNFiles", cdn.Private.Origins[0].Url);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_SiteDesigns()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var siteDesigns = hierarchy.Tenant.SiteDesigns;

            Assert.AreEqual("Just a sample", siteDesigns[0].Description);
            Assert.AreEqual(true, siteDesigns[0].IsDefault);
            Assert.AreEqual(false, siteDesigns[0].Overwrite);
            Assert.AreEqual("PnP Site Design Preview", siteDesigns[0].PreviewImageAltText);
            Assert.AreEqual("PnPSiteDesign.png", siteDesigns[0].PreviewImageUrl);
            Assert.AreEqual("PnP Site Design", siteDesigns[0].Title);
            Assert.AreEqual(1, (int)siteDesigns[0].WebTemplate); // TenantHelper.ProcessSiteDesigns handles conversion to SiteDesignWebTemplate from valid integer

            Assert.AreEqual("user1@contoso.com", siteDesigns[0].Grants[0].Principal);
            Assert.AreEqual(SiteDesignRight.View, siteDesigns[0].Grants[0].Right);
            Assert.AreEqual(SiteDesignRight.None, siteDesigns[0].Grants[2].Right);

            Assert.AreEqual("{SiteScriptID:PnP Site Script 01}", siteDesigns[0].SiteScripts[0]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Seserialize_Tenant_SiteDesigns()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new Core.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.SiteDesigns.Add(new SiteDesign
            {
                Description = "Just a sample",
                IsDefault = true,
                Overwrite = false,
                PreviewImageAltText = "PnP Site Design Preview",
                PreviewImageUrl = "PnPSiteDesign.png",
                Title = "PnP Site Design",
                WebTemplate = (SiteDesignWebTemplate)1, // xml serializer requires different numbers for web template
                Grants =
                {
                    new SiteDesignGrant
                    {
                        Principal = "user1@contoso.com",
                        Right = SiteDesignRight.View
                    }
                },
                SiteScripts = new List<string> { "{SiteScriptID:PnP Site Script 01}" }
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var siteDesigns = wrappedResult.Tenant.SiteDesigns;

            Assert.AreEqual("Just a sample", siteDesigns[0].Description);
            Assert.AreEqual(true, siteDesigns[0].IsDefault);
            Assert.AreEqual(false, siteDesigns[0].Overwrite);
            Assert.AreEqual("PnP Site Design Preview", siteDesigns[0].PreviewImageAltText);
            Assert.AreEqual("PnPSiteDesign.png", siteDesigns[0].PreviewImageUrl);
            Assert.AreEqual("PnP Site Design", siteDesigns[0].Title);
            Assert.AreEqual(SiteDesignsSiteDesignWebTemplate.CommunicationSite, siteDesigns[0].WebTemplate);

            Assert.AreEqual("user1@contoso.com", siteDesigns[0].Grants[0].Principal);
            Assert.AreEqual(SiteDesignsSiteDesignGrantRight.View, siteDesigns[0].Grants[0].Right);

            Assert.AreEqual("{SiteScriptID:PnP Site Script 01}", siteDesigns[0].SiteScripts[0].ID);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_SiteScripts()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");
            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var siteScripts = hierarchy.Tenant.SiteScripts;

            Assert.AreEqual("PnP Site Script Sample 01", siteScripts[0].Description);
            Assert.AreEqual(".\\pnp-site-script-01.json", siteScripts[0].JsonFilePath);
            Assert.AreEqual(true, siteScripts[0].Overwrite);
            Assert.AreEqual(false, siteScripts[1].Overwrite);
            Assert.AreEqual("PnP Site Script 01", siteScripts[0].Title);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Seserialize_Tenant_SiteScripts()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new Core.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.SiteScripts.Add(new SiteScript
            {
                Description = "PnP Site Script Sample 01",
                Overwrite = true,
                Title = "PnP Site Script 01",
                JsonFilePath = ".\\pnp-site-script-01.json"
            });

            result.Tenant.SiteScripts.Add(new SiteScript
            {
                Description = "PnP Site Script Sample 02",
                Overwrite = false,
                Title = "PnP Site Script 02",
                JsonFilePath = ".\\pnp-site-script-02.json"
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var siteScripts = wrappedResult.Tenant.SiteScripts;

            Assert.AreEqual("PnP Site Script Sample 01", siteScripts[0].Description);
            Assert.AreEqual(".\\pnp-site-script-01.json", siteScripts[0].JsonFilePath);
            Assert.AreEqual(true, siteScripts[0].Overwrite);
            Assert.AreEqual(false, siteScripts[1].Overwrite);
            Assert.AreEqual("PnP Site Script 01", siteScripts[0].Title);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_StorageEntities()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var storageEntities = hierarchy.Tenant.StorageEntities;

           Assert.AreEqual("Description 01", storageEntities[0].Description);
           Assert.AreEqual("Comment 01", storageEntities[0].Comment);
           Assert.AreEqual("PnPKey01", storageEntities[0].Key);
           Assert.AreEqual("My custom tenant-wide value 01", storageEntities[0].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Seserialize_Tenant_StorageEntities()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new Core.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.StorageEntities.Add(new StorageEntity
            {
                Description = "Description 01",
                Key = "PnPKey01",
                Value = "My custom tenant-wide value 01",
                Comment = "Comment 01"
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var storageEntities = wrappedResult.Tenant.StorageEntities;

            Assert.AreEqual("Description 01", storageEntities[0].Description);
            Assert.AreEqual("Comment 01", storageEntities[0].Comment);
            Assert.AreEqual("PnPKey01", storageEntities[0].Key);
            Assert.AreEqual("My custom tenant-wide value 01", storageEntities[0].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_Themes()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var themes = hierarchy.Tenant.Themes;

            Assert.AreEqual(false, themes[0].IsInverted);
            Assert.AreEqual("CustomOrange", themes[0].Name);
            Assert.IsTrue(themes[0].Palette.Contains("\"neutralQuaternaryAlt\": \"#dadada\""));

        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Seserialize_Tenant_Themes()
        {
            var provider = new XMLFileSystemTemplateProvider(String.Format(@"{0}\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new Core.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.Themes.Add(new Theme
            {
                Name = "CustomOrange",
                IsInverted = false,
                Palette = "{\"neutralQuaternaryAlt\": \"#dadada\"}"
            });

            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path));
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<Provisioning>(xml);
            var themes = wrappedResult.Tenant.Themes;

            Assert.AreEqual(false, themes[0].IsInverted);
            Assert.AreEqual("CustomOrange", themes[0].Name);
            Assert.IsTrue(themes[0].Text[0].Contains("\"neutralQuaternaryAlt\": \"#dadada\""));
        }
    }
}
