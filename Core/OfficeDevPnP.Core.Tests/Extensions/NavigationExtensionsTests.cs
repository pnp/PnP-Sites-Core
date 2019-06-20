using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Entities;

namespace OfficeDevPnP.Core.Tests.AppModelExtensions
{
    [TestClass]
    public class NavigationExtensionsTests
    {
        #region Add navigation node tests
        [TestMethod]
        public void AddTopNavigationNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                web.AddNavigationNode("Test Node", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.TopNavigationBar);

                clientContext.Load(web, w => w.Navigation.TopNavigationBar);
                clientContext.ExecuteQueryRetry();

                Assert.IsTrue(web.Navigation.TopNavigationBar.AreItemsAvailable);

                if (web.Navigation.TopNavigationBar.Any())
                {
                    var navNode = web.Navigation.TopNavigationBar.FirstOrDefault(n => n.Title == "Test Node");
                    Assert.IsNotNull(navNode);
                    navNode.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void AddQuickLaunchNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                web.AddNavigationNode("Test Node", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.QuickLaunch);

                clientContext.Load(web, w => w.Navigation.QuickLaunch);
                clientContext.ExecuteQueryRetry();

                Assert.IsTrue(web.Navigation.QuickLaunch.AreItemsAvailable);

                if (web.Navigation.QuickLaunch.Any())
                {
                    var navNode = web.Navigation.QuickLaunch.FirstOrDefault(n => n.Title == "Test Node");
                    Assert.IsNotNull(navNode);
                    navNode.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void AddSecondLevelQuickLaunchNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                web.AddNavigationNode("Level1", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.QuickLaunch);
                web.AddNavigationNode("Level2", new Uri("https://www.microsoft.com"), "Level1", NavigationType.QuickLaunch);

                clientContext.Load(web, w => w.Navigation.QuickLaunch);
                clientContext.ExecuteQueryRetry();

                Assert.IsTrue(web.Navigation.QuickLaunch.AreItemsAvailable);

                if (web.Navigation.QuickLaunch.Any())
                {
                    var l1NavNode = web.Navigation.QuickLaunch.FirstOrDefault(n => n.Title == "Level1");
                    Assert.IsNotNull(l1NavNode);

                    clientContext.Load(l1NavNode.Children);
                    clientContext.ExecuteQueryRetry();

                    var l2NavNode = l1NavNode.Children.FirstOrDefault(n => n.Title == "Level2");
                    Assert.IsNotNull(l2NavNode);

                    l2NavNode.DeleteObject();
                    clientContext.ExecuteQueryRetry();

                    l1NavNode.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void AddThirdLevelQuickLaunchNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                web.AddNavigationNode("Level1", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.QuickLaunch);
                web.AddNavigationNode("Level2", new Uri("https://www.microsoft.com"), "Level1", NavigationType.QuickLaunch);
                web.AddNavigationNode("Level3", new Uri("https://www.microsoft.com"), "Level2", NavigationType.QuickLaunch, l1ParentNodeTitle: "Level1");

                clientContext.Load(web, w => w.Navigation.QuickLaunch);
                clientContext.ExecuteQueryRetry();

                Assert.IsTrue(web.Navigation.QuickLaunch.AreItemsAvailable);

                if (web.Navigation.QuickLaunch.Any())
                {
                    var l1NavNode = web.Navigation.QuickLaunch.FirstOrDefault(n => n.Title == "Level1");
                    Assert.IsNotNull(l1NavNode);

                    clientContext.Load(l1NavNode.Children);
                    clientContext.ExecuteQueryRetry();

                    var l2NavNode = l1NavNode.Children.FirstOrDefault(n => n.Title == "Level2");
                    Assert.IsNotNull(l2NavNode);

                    clientContext.Load(l2NavNode.Children);
                    clientContext.ExecuteQueryRetry();

                    var l3NavNode = l2NavNode.Children.FirstOrDefault(n => n.Title == "Level3");
                    Assert.IsNotNull(l3NavNode);

                    l2NavNode.DeleteObject();
                    clientContext.ExecuteQueryRetry();

                    l1NavNode.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void AddSearchNavigationNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                // First clear all search nav nodes
                var nodeCollection = web.LoadSearchNavigation();
                foreach(var node in nodeCollection.ToList())
                {
                    node.DeleteObject();
                }
                clientContext.ExecuteQueryRetry();

                web.AddNavigationNode("Test Node", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.SearchNav);

                NavigationNodeCollection searchNavigation = web.LoadSearchNavigation();

                Assert.IsTrue(searchNavigation.AreItemsAvailable);

                if (searchNavigation.Any())
                {
                    var navNode = searchNavigation.FirstOrDefault(n => n.Title == "Test Node");
                    Assert.IsNotNull(navNode);
                    navNode.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }
        #endregion

        #region Delete navigation node tests
        [TestMethod]
        public void DeleteTopNavigationNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                web.AddNavigationNode("Test Node", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.TopNavigationBar);

                web.DeleteNavigationNode("Test Node", string.Empty, NavigationType.TopNavigationBar);

                clientContext.Load(web, w => w.Navigation.TopNavigationBar);
                clientContext.ExecuteQueryRetry();

                if (web.Navigation.TopNavigationBar.Any())
                {
                    var navNode = web.Navigation.TopNavigationBar.FirstOrDefault(n => n.Title == "Test Node");
                    Assert.IsNull(navNode);
                }
            }
        }

        [TestMethod]
        public void DeleteQuickLaunchNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                web.AddNavigationNode("Test Node", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.QuickLaunch);

                web.DeleteNavigationNode("Test Node", string.Empty, NavigationType.QuickLaunch);

                clientContext.Load(web, w => w.Navigation.QuickLaunch);
                clientContext.ExecuteQueryRetry();

                if (web.Navigation.QuickLaunch.Any())
                {
                    var navNode = web.Navigation.QuickLaunch.FirstOrDefault(n => n.Title == "Test Node");
                    Assert.IsNull(navNode);
                }
            }
        }

        [TestMethod]
        public void DeleteSearchNavigationNodeTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;

                // First clear all search nav nodes
                var nodeCollection = web.LoadSearchNavigation();
                foreach (var node in nodeCollection.ToList())
                {
                    node.DeleteObject();
                }
                clientContext.ExecuteQueryRetry();

                web.AddNavigationNode("Test Node", new Uri("https://www.microsoft.com"), string.Empty, NavigationType.SearchNav);

                web.DeleteNavigationNode("Test Node", string.Empty, NavigationType.SearchNav);

                NavigationNodeCollection searchNavigation = web.LoadSearchNavigation();

                if (searchNavigation.Any())
                {
                    var navNode = searchNavigation.FirstOrDefault(n => n.Title == "Test Node");
                    Assert.IsNull(navNode);
                }
            }
        }

        [TestMethod]
        public void DeleteAllNavigationNodesTest()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                var web = clientContext.Web;
                web.DeleteAllNavigationNodes(NavigationType.QuickLaunch);
                clientContext.Load(web, w => w.Navigation.QuickLaunch);
                clientContext.ExecuteQueryRetry();
                Assert.IsFalse(web.Navigation.QuickLaunch.Any());
            }
        }
        #endregion

    }
}
