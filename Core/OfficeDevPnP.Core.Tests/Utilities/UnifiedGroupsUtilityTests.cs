using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Graph;

namespace OfficeDevPnP.Core.Tests.Utilities
{
#if !ONPREMISES
    [TestClass]
    public class UnifiedGroupsUtilityTests
    {
        private string _accessToken;
        private string _groupId;

        #region Init and Cleanup code

        [TestInitialize]
        public void Initialize()
        {
            _accessToken = TestCommon.AcquireTokenAsync("https://graph.microsoft.com");

            if (!string.IsNullOrEmpty(_accessToken))
            {
                TestCommon.FixAssemblyResolving("Newtonsoft.Json");

                var random = new Random();
                _groupId = UnifiedGroupsUtility.CreateUnifiedGroup("PnPDeletedUnifiedGroup test", "PnPDeletedUnifiedGroup test", $"pnp-unit-test-{random.Next(1, 1000)}", _accessToken, groupLogo: null).GroupId;

                UnifiedGroupsUtility.DeleteUnifiedGroup(_groupId, _accessToken);
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            if (!string.IsNullOrEmpty(_accessToken))
            {
                try
                {
                    UnifiedGroupsUtility.DeleteUnifiedGroup(_groupId, _accessToken);
                }
                catch (Exception)
                {
                    // Group has already been deleted
                }
                try
                {
                    UnifiedGroupsUtility.PermanentlyDeleteUnifiedGroup(_groupId, _accessToken);
                }
                catch (Exception)
                {
                    // Group has already been permanently deleted
                }
            }
        }
        #endregion

        [TestMethod]
        public void ListDeletedUnifiedGroups()
        {
            if (string.IsNullOrEmpty(_accessToken)) Assert.Inconclusive("Access token could not be retrieved, so skipping this test");
            
            var results = UnifiedGroupsUtility.ListDeletedUnifiedGroups(_accessToken);

            Assert.IsTrue(results.Count > 0);
        }

        [TestMethod]
        public void GetDeletedUnifiedGroup()
        {
            if (string.IsNullOrEmpty(_accessToken)) Assert.Inconclusive("Access token could not be retrieved, so skipping this test");

            var results = UnifiedGroupsUtility.GetDeletedUnifiedGroup(_groupId, _accessToken);

            Assert.IsTrue(results != null && results.GroupId == _groupId);
        }

        [TestMethod]
        public void RestoreDeletedUnifiedGroup()
        {
            if (string.IsNullOrEmpty(_accessToken)) Assert.Inconclusive("Access token could not be retrieved, so skipping this test");

            UnifiedGroupsUtility.RestoreDeletedUnifiedGroup(_groupId, _accessToken);
            var results = UnifiedGroupsUtility.GetUnifiedGroup(_groupId, _accessToken);

            Assert.IsTrue(results != null && results.GroupId == _groupId);
        }

        [TestMethod]
        public void PermanentlyDeleteUnifiedGroup()
        {
            if (string.IsNullOrEmpty(_accessToken)) Assert.Inconclusive("Access token could not be retrieved, so skipping this test");

            UnifiedGroupsUtility.PermanentlyDeleteUnifiedGroup(_groupId, _accessToken);

            // The group should no longer be found in deleted groups
            try
            {
                var results = UnifiedGroupsUtility.GetDeletedUnifiedGroup(_groupId, _accessToken);
                Assert.IsFalse(results != null);
            }
            catch (Exception)
            {
                Assert.IsTrue(true);
            }
        }
    }
#endif
}
