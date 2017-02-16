using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class SecurityImplementation : ImplementationBase
    {
        internal void SiteCollectionSecurity(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.IncludeSiteGroups = true;
                ptci.HandlersToProcess = Handlers.SiteSecurity;

                var result = TestProvisioningTemplate(cc, "security_add.xml", Handlers.SiteSecurity, null, ptci);
                SecurityValidator sv = new SecurityValidator();
                Assert.IsTrue(sv.Validate(result.SourceTemplate.Security, result.TargetTemplate.Security, result.TargetTokenParser, cc));
            }
        }
    }
}