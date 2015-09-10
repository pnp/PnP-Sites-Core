using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Tests.Utilities
{
	[TestClass]
	public class UtilityTest
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
	}
}
