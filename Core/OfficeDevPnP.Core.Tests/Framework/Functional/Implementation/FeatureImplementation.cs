using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class FeatureImplementation: ImplementationBase
    {
        internal void SiteCollectionFeatureActivationDeactivation(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "feature_base.xml", Handlers.Features);
                Assert.IsTrue(FeatureValidator.Validate(result.SourceTemplate.Features, result.TargetTemplate.Features));
            }
        }

        internal void WebFeatureActivationDeactivation(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "feature_base.xml", Handlers.Features);
                Assert.IsTrue(FeatureValidator.ValidateFeatures(result.SourceTemplate.Features.WebFeatures, result.TargetTemplate.Features.WebFeatures));
            }
        }
    }
}
