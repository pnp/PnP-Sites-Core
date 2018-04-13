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
using System.Threading.Tasks;
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

        //[TestMethod]
        //public void BertTest5()
        //{
        //    using (var cc = TestCommon.CreateClientContext("https://bertonline.sharepoint.com/sites/ComSiteDemo"))
        //    {
        //        var page = cc.Web.LoadClientSidePage("home.aspx");

        //        page.Save("home2_normal.aspx");

        //    }
        //}


        //[TestMethod]
        //public void BertTest4()
        //{
        //    using (var cc = TestCommon.CreateClientContext("https://bertonline.sharepoint.com/sites/bert1"))
        //    {
        //        var newPage = new Pages.ClientSidePage(cc);
        //        //newPage.AddZone(CanvasZoneTemplate.OneColumn, 1);

        //        var imageWebPart = newPage.InstantiateDefaultWebPart(DefaultClientSideWebParts.Image);
        //        imageWebPart.PropertiesJson = "{\"id\":\"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788\",\"instanceId\":\"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788\",\"title\":\"Title Region\",\"description\":\"Title Region Description\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"imageSource\":\"/sites/bert1/Images1/1749309F-A8C2-4DD1-B3C3-42344AA58120.jpg\"},\"links\":{}},\"dataVersion\":\"1.3\",\"properties\":{\"title\":\"Imagedefaultfocalpoint\",\"textAlignCenter\":false,\"imageSourceType\":2,\"siteId\":\"78eaf8ed-fb6c-4bcb-a8ba-b4e251a90910\",\"webId\":\"ac56a969-5ca1-45fd-aca3-9ee5819e418f\",\"listId\":\"5d7a3301-0760-4472-97dd-af57f9cdd6f2\",\"uniqueId\":\"{D95285D8-7257-450E-88FC-6FCB3A03F42C}\"}}";
        //        newPage.AddControl(imageWebPart);


        //        //var t1 = new ClientSideText()
        //        //{
        //        //    Text = "t1"
        //        //};
        //        //var t2 = new ClientSideText()
        //        //{
        //        //    Text = "t2"
        //        //};

        //        //newPage.AddControl(t1, 0);
        //        //newPage.AddControl(t2, 1);
        //        //newPage.AddControl(t1, newPage.Zones[0].Sections[0], 2);
        //        //newPage.AddControl(t2, newPage.Zones[0].Sections[0], 1);

        //        newPage.Save("B1.aspx");

        //    }
        //}


        //[TestMethod]
        //public async Task GetAvailableClientSideComponentsTestAsync()
        //{
        //    using (var cc = TestCommon.CreateClientContext("https://bertonline.sharepoint.com/sites/bert1"))
        //    {
        //        var newPage = new Pages.ClientSidePage(cc);

        //        var components = await newPage.AvailableClientSideComponentsAsync("");

        //        Assert.IsTrue(components.Count() > 0);
        //    }
        //}

        #region Helper methods
        #endregion
    }
#endif
}
