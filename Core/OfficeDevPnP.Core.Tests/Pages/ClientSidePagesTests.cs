using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;
using System.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
#if !ONPREMISES
using OfficeDevPnP.Core.Pages;
#endif

namespace OfficeDevPnP.Core.Tests.Authentication
{
#if !ONPREMISES
    [TestClass]
    public class ClientSidePagesTests
    {

#region Test initialization
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {

        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                
            }
        }
        #endregion

        //[TestMethod]
        //public void PageTest()
        //{
        //    using (var clientContext = TestCommon.CreateClientContext("https://bertonline.sharepoint.com/sites/bert1"))
        //    {
        //        ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(clientContext.Web)
        //        {
        //            // Limit the amount of handlers in this demo
        //            HandlersToProcess = Handlers.PageContents,
        //            // Create FileSystemConnector, so that we can store composed files temporarely somewhere 
        //            FileConnector = new FileSystemConnector(@"C:\temp", ""),
        //            //PersistBrandingFiles = true,
        //            ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
        //            {
        //                // Only to output progress for console UI
        //                Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
        //            }
        //        };

        //        // Execute actual extraction of the tepmplate
        //        ProvisioningTemplate template = clientContext.Web.GetProvisioningTemplate(ptci);

        //        // Serialize to XML using the beta 201705 schema
        //        XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(@"C:\temp", "");
        //        var formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05);
        //        provider.SaveAs(template, "PnPProvisioningDemo201705.xml", formatter);
        //    }
        //}

        //[TestMethod]
        //public void Page2Test()
        //{
        //    using (var clientContext = TestCommon.CreateClientContext("https://bertonline.sharepoint.com/sites/bert1"))
        //    {
        //        ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation()
        //        {
        //            HandlersToProcess = Handlers.PageContents,
        //            ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
        //            {
        //                    // Only to output progress for console UI
        //                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
        //            }
        //        };

        //        XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(@"c:\temp", "");
        //        ProvisioningTemplate sourceTemplate = provider.GetTemplate("PnPProvisioningDemo201705_load.xml");

        //        // Execute actual extraction of the tepmplate
        //        clientContext.Web.ApplyProvisioningTemplate(sourceTemplate);
        //    }
        //}
        #region Helper methods
        #endregion
    }
#endif
}
