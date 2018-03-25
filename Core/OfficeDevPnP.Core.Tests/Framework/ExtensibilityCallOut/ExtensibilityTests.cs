using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

#if !NETSTANDARD2_0
namespace OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut
{
    [TestClass]
    public class ExtensibilityTests
    {
        private const string TEST_CATEGORY = "Framework Provisioning Extensibility Providers";

#region Providers

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void CanProviderCallOut()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "OfficeDevPnP.Core.Tests";
            _mockProvider.Type = "OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityMockProvider";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteExtensibilityCallOut(_mockctx, _mockProvider, _mockTemplate);
        }


        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void CanHandlerProvisioningCallOut()
        {
            var _mockProvider = new ExtensibilityHandler();
            _mockProvider.Assembly = "OfficeDevPnP.Core.Tests";
            _mockProvider.Type = "OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityMockHandler";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteExtensibilityProvisionCallOut(_mockctx, _mockProvider, _mockTemplate, null, null, null);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void CanHandlerExtractionCallOut()
        {
            var _mockProvider = new ExtensibilityHandler();
            _mockProvider.Assembly = "OfficeDevPnP.Core.Tests";
            _mockProvider.Type = "OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityMockHandler";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            var template = _em.ExecuteExtensibilityExtractionCallOut(_mockctx, _mockProvider, _mockTemplate, null, null);

            Assert.IsTrue(template.Lists.Count == 1);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ExtensiblityPipelineException))]
        public void ProviderCallOutThrowsException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "BLAHASSEMLBY";
            _mockProvider.Type = "BLAHTYPE";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteExtensibilityCallOut(_mockctx, _mockProvider, _mockTemplate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ArgumentException))]
        public void ProviderAssemblyMissingThrowsAgrumentException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "";
            _mockProvider.Type = "TYPE";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteExtensibilityCallOut(_mockctx, _mockProvider, _mockTemplate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ArgumentException))]
        public void ProviderTypeNameMissingThrowsAgrumentException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "BLAHASSEMBLY";
            _mockProvider.Type = "";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteExtensibilityCallOut(_mockctx, _mockProvider, _mockTemplate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ProviderClientCtxIsNullThrowsAgrumentNullException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "BLAHASSEMBLY";
            _mockProvider.Type = "BLAH";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            ClientContext _mockCtx = null;
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteExtensibilityCallOut(_mockCtx, _mockProvider, _mockTemplate);
        }

#endregion

#region TokenProviders

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void TokenProviderReceivesExpectedParameters()
        {
            var givenConfiguration = "START {parameter:MOCKPARAM} END";
            var expectedConfiguration = "START MOCKPARAMVALUE END";

            using (var ctx = TestCommon.CreateClientContext())
            {
                var _mockProvider = new Provider
                {
                    Assembly = "OfficeDevPnP.Core.Tests",
                    Type = "OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityMockTokenProvider",
                    Configuration = givenConfiguration,
                    Enabled = true
                };

                var provisioningInfo = new ProvisioningTemplateApplyingInformation();
                provisioningInfo.HandlersToProcess = Handlers.All;
                provisioningInfo.ExtensibilityHandlers.Add(_mockProvider);

                var _mockTemplate = new ProvisioningTemplate();
                _mockTemplate.Parameters.Add("MOCKPARAM", "MOCKPARAMVALUE");
                _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;
                _mockTemplate.ExtensibilityHandlers.Add(_mockProvider);

                var extensibilityHandler = new ObjectExtensibilityHandlers();
                var parser = new TokenParser(ctx.Web, _mockTemplate);
                extensibilityHandler.AddExtendedTokens(ctx.Web, _mockTemplate, parser, provisioningInfo);

                Assert.AreSame(ctx, ExtensibilityMockTokenProvider.ReceivedCtx, "Wrong clientContext passed to the provider.");
                Assert.AreSame(_mockTemplate, ExtensibilityMockTokenProvider.ReceivedProvisioningTemplate, "Wrong template passed to the provider.");
                Assert.AreEqual(expectedConfiguration, ExtensibilityMockTokenProvider.ReceivedConfigurationData, "Wrong configuration data passed to the provider.");
            }
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void TokenProviderProvidesTokens()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var _mockProvider = new Provider
                {
                    Assembly = "OfficeDevPnP.Core.Tests",
                    Type = "OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityMockTokenProvider",
                    Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA,
                    Enabled = true
                };

                var provisioningInfo = new ProvisioningTemplateApplyingInformation();
                provisioningInfo.HandlersToProcess = Handlers.All;
                provisioningInfo.ExtensibilityHandlers.Add(_mockProvider);

                var _mockTemplate = new ProvisioningTemplate();
                _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;
                _mockTemplate.ExtensibilityHandlers.Add(_mockProvider);

                var extensibilityHandler = new ObjectExtensibilityHandlers();
                var parser = new TokenParser(ctx.Web, _mockTemplate);
                extensibilityHandler.AddExtendedTokens(ctx.Web, _mockTemplate, parser, provisioningInfo);

                var parsedValue = parser.ParseString(MockToken.MockTokenKey);
                Assert.AreEqual(MockToken.MockTokenValue, parsedValue);
            }
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void TokenProviderCanBeDisabled()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var _mockProvider = new Provider
                {
                    Assembly = "OfficeDevPnP.Core.Tests",
                    Type = "OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut.ExtensibilityMockTokenProvider",
                    Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA,
                    Enabled = false
                };

                var provisioningInfo = new ProvisioningTemplateApplyingInformation();
                provisioningInfo.HandlersToProcess = Handlers.All;
                provisioningInfo.ExtensibilityHandlers.Add(_mockProvider);

                var _mockTemplate = new ProvisioningTemplate();
                _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;
                _mockTemplate.ExtensibilityHandlers.Add(_mockProvider);

                var extensibilityHandler = new ObjectExtensibilityHandlers();
                var parser = new TokenParser(ctx.Web, _mockTemplate);
                extensibilityHandler.AddExtendedTokens(ctx.Web, _mockTemplate, parser, provisioningInfo);

                var parsedValue = parser.ParseString(MockToken.MockTokenKey);
                Assert.AreEqual(MockToken.MockTokenKey, parsedValue, "Disabled tokenprovider should not have provided tokens!");
            }
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ExtensiblityPipelineException))]
        public void TokenProviderCallOutThrowsException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "BLAHASSEMLBY";
            _mockProvider.Type = "BLAHTYPE";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteTokenProviderCallOut(_mockctx, _mockProvider, _mockTemplate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ArgumentException))]
        public void TokenProviderAssemblyMissingThrowsAgrumentException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "";
            _mockProvider.Type = "TYPE";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteTokenProviderCallOut(_mockctx, _mockProvider, _mockTemplate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ArgumentException))]
        public void TokenProviderTypeNameMissingThrowsAgrumentException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "BLAHASSEMBLY";
            _mockProvider.Type = "";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            var _mockctx = new ClientContext(ExtensibilityTestConstants.MOCK_URL);
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteTokenProviderCallOut(_mockctx, _mockProvider, _mockTemplate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        [ExpectedException(typeof(ArgumentNullException))]
        public void TokenProviderClientCtxIsNullThrowsAgrumentNullException()
        {
            var _mockProvider = new Provider();
            _mockProvider.Assembly = "BLAHASSEMBLY";
            _mockProvider.Type = "BLAH";
            _mockProvider.Configuration = ExtensibilityTestConstants.PROVIDER_MOCK_DATA;

            ClientContext _mockCtx = null;
            var _mockTemplate = new ProvisioningTemplate();
            _mockTemplate.Id = ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID;

            var _em = new ExtensibilityManager();
            _em.ExecuteTokenProviderCallOut(_mockCtx, _mockProvider, _mockTemplate);
        }

#endregion
    }
}
#endif