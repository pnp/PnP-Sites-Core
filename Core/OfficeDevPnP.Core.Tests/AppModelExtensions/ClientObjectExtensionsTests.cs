using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.AppModelExtensions
{
    [TestClass()]
    public class ClientObjectExtensionsTests
    {
        [TestMethod]
        public void NotLoadedPropertyExceptionTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                try
                {
                    //Act
                    var webUrl = clientContext.Web.ServerRelativeUrl;
                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsTrue(expectedException is PropertyOrFieldNotInitializedException);
            }
        }

        [TestMethod]
        public void EnsurePropertyTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                try
                {
                    //Act
                    var serverRelativeUrl = clientContext.Web.EnsureProperty(w => w.ServerRelativeUrl);
                    var id = clientContext.Web.EnsureProperty(w => w.Id);
                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsNull(expectedException);
                Assert.IsTrue(clientContext.Web.IsPropertyAvailable(w => w.ServerRelativeUrl));
            }
        }

        [TestMethod]
        public void NotLoadedCollectionExceptionTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                try
                {
                    //Act
                    var webFields = clientContext.Web.Fields.ToList();
                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsTrue(expectedException is CollectionNotInitializedException);
            }
        }

        [TestMethod]
        public void EnsureCollectionPropertyTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                try
                {
                    //Act
                    var webFields = clientContext.Web.EnsureProperty(w => w.Fields).ToList();
                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsNull(expectedException);
                Assert.IsTrue(clientContext.Web.IsObjectPropertyInstantiated(w => w.Fields));
            }
        }

        [TestMethod]
        public void NotLoadedComplexPropertyExceptionTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                try
                {
                    //Act
                    var rootFolderUrl = clientContext.Web.RootFolder.ServerRelativeUrl;
                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsTrue(expectedException is PropertyOrFieldNotInitializedException);
            }
        }

        [TestMethod]
        public void EnsureComplexPropertyTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                try
                {
                    //Act
                    var rootFolderUrl = clientContext.Web.EnsureProperty(f => f.RootFolder);

                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsNull(expectedException);
                Assert.IsTrue(clientContext.Web.IsObjectPropertyInstantiated(w => w.RootFolder));
                Assert.IsNotNull(clientContext.Web.RootFolder.ServerRelativeUrl);
            }
        }

        [TestMethod]
        public void EnsureMultiplePropertiesTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                try
                {
                    //Act
                    clientContext.Web.EnsureProperties(w => w.Fields, w => w.ServerRelativeUrl);

                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsNull(expectedException);
                Assert.IsTrue(clientContext.Web.IsObjectPropertyInstantiated(w => w.Fields));
                Assert.IsTrue(clientContext.Web.IsPropertyAvailable(w => w.ServerRelativeUrl));
            }
        }

        [TestMethod]
        public void EnsurePropertiesIncludeTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                Field field = null;
                string url = null;
                try
                {
                    //Act
                    clientContext.Web.EnsureProperties(w => w.Fields.Include(f => f.Id, f => f.Title), w => w.Url);

                    //equivalent to
                    //clientContext.Load(clientContext.Web, w=> w.Url,  w => w.Fields.Include(f => f.Id, f => f.Title));
                    //clientContext.ExecuteQueryRetry();

                    field = clientContext.Web.Fields[0];
                    url = clientContext.Web.Url;
                    var hidden = field.Required;
                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsTrue(expectedException is PropertyOrFieldNotInitializedException);
                Assert.IsTrue(!string.IsNullOrEmpty(field.Title));
                Assert.IsTrue(!string.IsNullOrEmpty(url));
            }
        }

        [TestMethod]
        public void EnsurePropertyIncludeTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                Exception expectedException = null;
                Field field = null;
                try
                {
                    //Act
                    var fields = clientContext.Web.EnsureProperty(w => w.Fields.Include(f => f.Id, f => f.Title)).ToList();
                    
                    //equivalent to
                    //clientContext.Load(clientContext.Web, w => w.Fields.Include(f => f.Id, f => f.Title));
                    //clientContext.ExecuteQueryRetry();
                    //var fields = clientContext.Web.Fields;

                    field = fields[0];
                    
                    var hidden = field.Required;
                }
                catch (Exception ex)
                {
                    expectedException = ex;
                }

                //Assert
                Assert.IsTrue(expectedException is PropertyOrFieldNotInitializedException);
                Assert.IsTrue(!string.IsNullOrEmpty(field.Title));
            }
        }

        [TestMethod]
        public void EnsureComplexPropertyWithDependencyTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                var fieldTitle1 = clientContext.Web.Fields.GetByInternalNameOrTitle("Title");
                //at this stage clientContext.Web.IsObjectPropertyInstantiated("Fields") will be true
                //but actually Fields are not loaded, CollectionNotInitializedException will be thrown when trying to access the collection
                
                var fields = clientContext.Web.EnsureProperty(w => w.Fields);
                
                //Act
                var fieldTitle2 = fields.FirstOrDefault(f => f.Title.Equals("Title"));

                //Assert
                Assert.IsTrue(fieldTitle2 != null);
                Assert.IsTrue(fieldTitle1 != null);
            }
        }

        [TestMethod]
        public void EnsureComplexPropertiesWithDependencyTest()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                //Arrange
                var fieldTitle1 = clientContext.Web.Fields.GetByInternalNameOrTitle("Title");
                //at this stage clientContext.Web.IsObjectPropertyInstantiated("Fields") will be true
                //but actually Fields are not loaded, CollectionNotInitializedException will be thrown when trying to access the collection

                clientContext.Web.EnsureProperties(w => w.Fields);

                //Act
                var fieldTitle2 = clientContext.Web.Fields.FirstOrDefault(f => f.Title.Equals("Title"));

                //Assert
                Assert.IsTrue(fieldTitle2 != null);
                Assert.IsTrue(fieldTitle1 != null);
            }
        }
    }
}
