using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectComposedLookTests
    {

        [TestMethod]
        public void CanCreateComposedLooks()
        {
            using (var scope = new Core.Diagnostics.PnPMonitoredScope("ComposedLookTests"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                    var template = new ProvisioningTemplate();
                    template = new ObjectComposedLook().ExtractObjects(ctx.Web, template, creationInfo);
                    Assert.IsInstanceOfType(template.ComposedLook, typeof(Core.Framework.Provisioning.Model.ComposedLook));
                }
            }
        }
    }
}
