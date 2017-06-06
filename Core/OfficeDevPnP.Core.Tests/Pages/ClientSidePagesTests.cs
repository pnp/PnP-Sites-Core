using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;
using System.Linq;
using OfficeDevPnP.Core.Pages;

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
                //DeleteListsImplementation(clientContext);
            }
        }
        #endregion

        [TestMethod]
        public void BertTest()
        {
            using (var cc = TestCommon.CreateClientContext("https://a830edad9050849523e17050400.sharepoint.com/sites/modern"))
            {
                //var page = cc.Web.LoadClientSidePage("Demo1.aspx");
                //var newHomePage = cc.Web.AddClientSidePage();
                //newHomePage.LayoutType = ClientSidePageLayoutType.Home;
                //newHomePage.KeepDefaultWebParts = true;
                //newHomePage.Save("BertHome2.aspx");
                //newHomePage.Publish();

                var newPage = cc.Web.AddClientSidePage();

                //newPage.AddZone(new CanvasZone(newPage, CanvasZoneTemplate.TwoColumnLeft, 1));
                //newPage.AddZone(new CanvasZone(newPage, CanvasZoneTemplate.TwoColumnRight, 2));
                newPage.AddZone(CanvasZoneTemplate.TwoColumnLeft, 1);
                newPage.AddZone(CanvasZoneTemplate.TwoColumnRight, 2);

                var t1 = new ClientSideText()
                {
                    Text = "t1"
                };
                var t2 = new ClientSideText()
                {
                    Text = "t2"
                };

                newPage.AddControl(t1, newPage.Zones[0].Sections[1]);
                newPage.AddControl(t2, newPage.Zones[1].Sections[0]);

                newPage.Save("B4.aspx");
                newPage.Publish();

            }
        }

        [TestMethod]
        public void BertTest2()
        {
            using (var cc = TestCommon.CreateClientContext("https://a830edad9050849523e17050400.sharepoint.com/sites/modern"))
            {
                var newPage = cc.Web.AddClientSidePage();
                newPage.AddZone(CanvasZoneTemplate.OneColumnFullWidth, 1);

                var t1 = new ClientSideText()
                {
                    Text = "t1"
                };
                var t2 = new ClientSideText()
                {
                    Text = "t2"
                };

                newPage.AddControl(t1, newPage.Zones[0].Sections[0], 3);
                newPage.AddControl(t2, newPage.Zones[0].Sections[0], 1);

                var people = newPage.AvailableClientSideComponents(DefaultClientSideWebParts.People);
                var c1 = new ClientSideWebPart(people.First());
                newPage.AddControl(c1, newPage.Zones[0].Sections[0], 2);

                newPage.Save("B6.aspx");
            }
        }

        [TestMethod]
        public void BertTest3()
        {
            using (var cc = TestCommon.CreateClientContext("https://a830edad9050849523e17050400.sharepoint.com/sites/modern"))
            {
                var page = cc.Web.LoadClientSidePage("berthome2.aspx");
                //page.DemoteNewsArticle();
                //page.PromoteAsNewsArticle();
                //page.Publish();
                //page.PromoteAsHomePage();

                var t1 = new ClientSideText()
                {
                    Text = "t1"
                };
                var t2 = new ClientSideText()
                {
                    Text = "t2"
                };

                page.Controls[0].Order = 5;
                page.Controls[1].Order = 7;
                page.AddControl(t1, page.Zones[0].Sections[0], 0);
                page.AddControl(t2, page.Zones[0].Sections[0], 10);

                //page.KeepDefaultWebParts = true;
                page.Save();

            }
        }

        [TestMethod]
        public void BertTest5()
        {
            using (var cc = TestCommon.CreateClientContext("https://a830edad9050849523e17050400.sharepoint.com/sites/modern"))
            {
                var page = cc.Web.LoadClientSidePage("b6.aspx");
                //page.DemoteNewsArticle();
                //page.PromoteAsNewsArticle();
                //page.Publish();
                //page.PromoteAsHomePage();
                var commentsDisabled = page.CommentsDisabled;
                page.DisableComments();

            }
        }

        [TestMethod]
        public void BertTest4()
        {
            using (var cc = TestCommon.CreateClientContext("https://bertonline.sharepoint.com/sites/uiscannermodernteam"))
            {
                var newPage = new ClientSidePage(cc);
                //newPage.AddZone(CanvasZoneTemplate.OneColumn, 1);

                var imageWebPart = newPage.InstantiateDefaultWebPart(DefaultClientSideWebParts.Image);
                imageWebPart.Properties["imageSourceType"] = 2;
                imageWebPart.Properties["siteId"] = "c827cb03-d059-4956-83d0-cd60e02e3b41";
                imageWebPart.Properties["webId"] = "9fafd7c0-e8c3-4a3c-9e87-4232c481ca26";
                imageWebPart.Properties["listId"] = "78d1b1ac-7590-49e7-b812-55f37c018c4b";
                imageWebPart.Properties["uniqueId"] = "3C27A419-66D0-4C36-BF24-BD6147719052";
                imageWebPart.Properties["imgWidth"] = 1002;
                imageWebPart.Properties["imgHeight"] = 469;
                newPage.AddControl(imageWebPart);


                //var t1 = new ClientSideText()
                //{
                //    Text = "t1"
                //};
                //var t2 = new ClientSideText()
                //{
                //    Text = "t2"
                //};

                //newPage.AddControl(t1, 0);
                //newPage.AddControl(t2, 1);
                //newPage.AddControl(t1, newPage.Zones[0].Sections[0], 2);
                //newPage.AddControl(t2, newPage.Zones[0].Sections[0], 1);

                newPage.Save("B1.aspx");

            }
        }

        #region Helper methods
        #endregion
    }
#endif
}
