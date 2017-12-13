using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectSupportedUILanguagesTests
    {
        [TestMethod]
        public void CanExtractSupportedUILanguages()
        {
            using (var scope = new Core.Diagnostics.PnPMonitoredScope("CanProvisionSupportedUILanguages"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = null };

                    var template = new ProvisioningTemplate();
                    template = new ObjectSupportedUILanguages().ExtractObjects(ctx.Web, template, creationInfo);

                    Assert.IsTrue(template.SupportedUILanguages.Count > 0);

                }
            }
        }

        [TestMethod]
        public void CanProvisionSupportedUILanguages()
        {
            using (var scope = new Core.Diagnostics.PnPMonitoredScope("CanProvisionSupportedUILanguages"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    var template = new ProvisioningTemplate();

                    template.SupportedUILanguages.Add(new SupportedUILanguage() { LCID = 1033 }); // English
                    template.SupportedUILanguages.Add(new SupportedUILanguage() { LCID = 1032 }); // Greek

                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectSupportedUILanguages().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                    ctx.Load(ctx.Web, w => w.SupportedUILanguageIds);

                    ctx.ExecuteQueryRetry();

                    Assert.IsTrue(ctx.Web.SupportedUILanguageIds.Count() == 2);
                    Assert.IsTrue(ctx.Web.SupportedUILanguageIds.Any(i => i == 1033));
                    Assert.IsTrue(ctx.Web.SupportedUILanguageIds.Any(i => i == 1032));
                }
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Load(ctx.Web, w => w.Language);
                ctx.Load(ctx.Web, w => w.SupportedUILanguageIds);
                ctx.ExecuteQueryRetry();

                foreach (var id in ctx.Web.SupportedUILanguageIds)
                {
                    if (id != ctx.Web.Language)
                    {
                        ctx.Web.RemoveSupportedUILanguage(id);
                    }
                }
                ctx.Web.Update();
                ctx.Web.IsMultilingual = false;
                ctx.ExecuteQueryRetry();
            }
        }
    }
}
