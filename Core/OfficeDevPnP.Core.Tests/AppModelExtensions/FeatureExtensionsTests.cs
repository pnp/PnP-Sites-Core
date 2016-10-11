using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Tests;
namespace Microsoft.SharePoint.Client.Tests
{
    [TestClass()]
    public class FeatureExtensionsTests
    {
        const string TEST_CATEGORY = "Feature Extensions";
        private ClientContext clientContext;
        private Guid publishingSiteFeatureId = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");
        private Guid publishingWebFeatureId = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
        private Guid sp2007WorkflowSiteFeatureId = new Guid("c845ed8d-9ce5-448c-bd3e-ea71350ce45b");
        private Guid contentOrganizerWebFeatureId = new Guid("7ad5272a-2694-4349-953e-ea5ef290e97c");

        #region Test initialize and cleanup
        [TestInitialize()]
        public void Initialize()
        {
            clientContext = TestCommon.CreateClientContext();
        }

        [TestCleanup()]
        public void Cleanup()
        {
            clientContext.Dispose();
        }
        #endregion

        #region Feature activation tests
        [TestMethod()]
        public void ActivateSiteFeatureTest()
        {
            // Test
            clientContext.Site.ActivateFeature(sp2007WorkflowSiteFeatureId);

            Assert.IsTrue(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));

            clientContext.Site.DeactivateFeature(sp2007WorkflowSiteFeatureId);
            
            Assert.IsFalse(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));
        }

        [TestMethod()]
        public void ActivateWebFeatureTest()
        {
            // Test
            clientContext.Web.ActivateFeature(contentOrganizerWebFeatureId);

            Assert.IsTrue(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));

            clientContext.Web.DeactivateFeature(contentOrganizerWebFeatureId);

            Assert.IsFalse(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));
        }

        [TestMethod()]
        public void PublishingFeatureTest()
        {
            // First deactivate if it was still active
            clientContext.Web.DeactivateFeature(publishingWebFeatureId);
            Assert.IsFalse(clientContext.Web.IsFeatureActive(publishingWebFeatureId));

            clientContext.Site.DeactivateFeature(publishingSiteFeatureId);
            Assert.IsFalse(clientContext.Site.IsFeatureActive(publishingSiteFeatureId));

            // Activate
            clientContext.Site.ActivateFeature(publishingSiteFeatureId);
            Assert.IsTrue(clientContext.Site.IsFeatureActive(publishingSiteFeatureId));

            clientContext.Web.ActivateFeature(publishingWebFeatureId);
            Assert.IsTrue(clientContext.Web.IsFeatureActive(publishingWebFeatureId));

            // Finally deactivate again
            clientContext.Web.DeactivateFeature(publishingWebFeatureId);
            Assert.IsFalse(clientContext.Web.IsFeatureActive(publishingWebFeatureId));

            clientContext.Site.DeactivateFeature(publishingSiteFeatureId);
            Assert.IsFalse(clientContext.Site.IsFeatureActive(publishingSiteFeatureId));
        }

        [TestMethod()]
        public void DeactivateSiteFeatureTest()
        {
            // Setup
            clientContext.Site.ActivateFeature(sp2007WorkflowSiteFeatureId);

            // Test
            clientContext.Site.DeactivateFeature(sp2007WorkflowSiteFeatureId);
            Assert.IsFalse(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));
        }

        [TestMethod()]
        public void DeactivateWebFeatureTest()
        {
            // Setup
            clientContext.Web.ActivateFeature(contentOrganizerWebFeatureId);

            // Test
            clientContext.Web.DeactivateFeature(contentOrganizerWebFeatureId);
            Assert.IsFalse(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));
        }

        [TestMethod()]
        public void IsSiteFeatureActiveTest()
        {
            // Setup
            try
            {
                clientContext.Site.DeactivateFeature(sp2007WorkflowSiteFeatureId);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ignoring exception: {0}", ex.Message);
            }

            // Test
            Assert.IsFalse(clientContext.Site.IsFeatureActive(sp2007WorkflowSiteFeatureId));
        }

        [TestMethod()]
        public void IsWebFeatureActiveTest()
        {
            // Setup
            try
            { 
                clientContext.Web.DeactivateFeature(contentOrganizerWebFeatureId);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ignoring exception: {0}", ex.Message);
            }

            // Test
            Assert.IsFalse(clientContext.Web.IsFeatureActive(contentOrganizerWebFeatureId));
        }
        #endregion
    }
}
