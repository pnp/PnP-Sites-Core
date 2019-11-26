#if !SP2013 && !SP2016
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Tests.Utilities;
using OfficeDevPnP.Core.Utilities;
using System;

namespace OfficeDevPnP.Core.Tests.Framework.ProvisioningTemplates
{
    [TestClass]
    public class ClientSidePageProvisioningTests
    {

        [TestCleanup]
        public void Cleanup()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();

                TestCommon.DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/csp-test-1.aspx"));
                TestCommon.DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/csp-test-2.aspx"));
                TestCommon.DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/csp-test-3.aspx"));
            }
        }

        // background for this test: https://github.com/SharePoint/PnP-Sites-Core/issues/2203
        [TestMethod]
        public void ProvisionClientSidePagesWithHeader()
        {
            var resourceFolder = string.Format(@"{0}\Resources\Templates", AppDomain.CurrentDomain.BaseDirectory);
            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(resourceFolder, "");

            var existingTemplate = provider.GetTemplate("ClientSidePagesWithHeader.xml");
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Web.ApplyProvisioningTemplate(existingTemplate, new ProvisioningTemplateApplyingInformation()
                {
                    HandlersToProcess = Handlers.Pages
                });
            }
        }
    }
}
#endif