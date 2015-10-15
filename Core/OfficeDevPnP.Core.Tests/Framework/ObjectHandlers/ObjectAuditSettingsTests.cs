using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectAuditSettingsTests
    {
        private Audit audit;
        private int auditLogTrimmingRetention;
        private bool trimAuditLog;

        [TestInitialize]
        public void Initialize()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var site = ctx.Site;
                audit = site.Audit;
                ctx.Load(audit, af => af.AuditFlags);
                ctx.Load(site, s => s.AuditLogTrimmingRetention, s => s.TrimAuditLog);
                ctx.ExecuteQueryRetry();

                auditLogTrimmingRetention = site.AuditLogTrimmingRetention;
                trimAuditLog = site.TrimAuditLog;

                // set audit flags for test
                site.Audit.AuditFlags = AuditMaskType.All;
                site.Audit.Update();
                ctx.ExecuteQueryRetry();
            }
        }

        [TestMethod]
        public void CanExtractAuditSettings()
        {
            using (var scope = new Core.Diagnostics.PnPMonitoredScope("CanExtractAuditSettings"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };
                    var template = new ProvisioningTemplate();
                    template = new ObjectAuditSettings().ExtractObjects(ctx.Web, template, creationInfo);

                    Assert.IsNotNull(template.AuditSettings);



                }
            }
        }

        [TestMethod]
        public void CanProvisionAuditSettings()
        {
            using (var scope = new Core.Diagnostics.PnPMonitoredScope("CanProvisionAuditSettings"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    var template = new ProvisioningTemplate();

                    template.AuditSettings = new AuditSettings();
                    template.AuditSettings.AuditFlags = AuditMaskType.CheckIn;
                    template.AuditSettings.AuditLogTrimmingRetention = 5;
                    template.AuditSettings.TrimAuditLog = true;

                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectAuditSettings().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                    var site = ctx.Site;

                    var auditSettings = site.Audit;
                    ctx.Load(auditSettings, af => af.AuditFlags);
                    ctx.Load(site, s => s.AuditLogTrimmingRetention, s => s.TrimAuditLog);
                    ctx.ExecuteQueryRetry();

                    Assert.IsTrue(auditSettings.AuditFlags == AuditMaskType.CheckIn);
                    Assert.IsTrue(site.AuditLogTrimmingRetention == 5);
                    Assert.IsTrue(site.TrimAuditLog = true);
                }
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var site = ctx.Site;
                site.Audit.AuditFlags = audit.AuditFlags;
                site.Audit.Update();
                site.AuditLogTrimmingRetention = auditLogTrimmingRetention;
                site.TrimAuditLog = trimAuditLog;
                ctx.ExecuteQueryRetry();
            }
        }
    }
}
