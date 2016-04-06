using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Tests.Framework.Model
{
    [TestClass]
    public class CustomActionTests
    {
        [TestMethod]
        public void SetBasePermissionsTest()
        {
            CustomAction ca = new CustomAction();
            BasePermissions bp = new BasePermissions();
            bp.Set(PermissionKind.ApplyStyleSheets);
            bp.Set(PermissionKind.BrowseUserInfo);

            ca.Rights = bp;

            ca.RightsValue = 29;

        }
    }
}
