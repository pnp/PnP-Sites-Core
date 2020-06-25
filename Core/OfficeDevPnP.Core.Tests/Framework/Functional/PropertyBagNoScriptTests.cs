using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !SP2013 && !SP2016
    [TestClass]
    public class PropertyBagNoScriptTests: FunctionalTestBase
    {

        #region Construction
        public PropertyBagNoScriptTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_25b60217-025d-45a8-961c-7436cb7419df";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_25b60217-025d-45a8-961c-7436cb7419df/sub";
        }
        #endregion

        #region Test setup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            ClassInitBase(context, true);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            ClassCleanupBase();
        }
        #endregion

        #region Site collection test cases
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionPropertyBagAddingTest()
        {
            new PropertyBagImplementation().SiteCollectionPropertyBagAdding(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebPropertyBagAddingTest()
        {
            new PropertyBagImplementation().WebPropertyBagAdding(centralSubSiteUrl);
        }
        #endregion
    }
#endif
}
