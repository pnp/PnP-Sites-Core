using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Utilities.WebParts;
using OfficeDevPnP.Core.Utilities.WebParts.Processors;

namespace OfficeDevPnP.Core.Tests.Utilities.WebParts
{
    [TestClass]
    public class WebPartPostProcessorFactoryTests
    {
        [TestMethod]
        public void ReturnsPassThroughProcessorForUknownSchema()
        {
            string wpXml = @"
                    <webParts>
                        <webPart>
                            <metaData></metaData>
                        <data>
                          <properties>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            var postProcessor = WebPartPostProcessorFactory.Resolve(wpXml);

            Assert.IsTrue(postProcessor is PassThroughProcessor);
        }

        [TestMethod]
        public void ReturnsPassThroughProcessorForUndefinedType()
        {
            string wpXml = @"
                    <webParts>
                        <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                            <metaData><type name=""Uknown"" /></metaData>
                        <data>
                          <properties>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            var postProcessor = WebPartPostProcessorFactory.Resolve(wpXml);

            Assert.IsTrue(postProcessor is PassThroughProcessor);
        }

        [TestMethod]
        public void ReturnsPassThroughProcessorForUknownWebPartType()
        {
            string wpXml = @"
                    <webParts>
                        <webPart>
                            <metaData><type name=""Uknown"" /></metaData>
                        <data>
                          <properties>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            var postProcessor = WebPartPostProcessorFactory.Resolve(wpXml);

            Assert.IsTrue(postProcessor is PassThroughProcessor);
        }

        [TestMethod]
        public void ReturnsXsltWebPartPostProcessor()
        {
            string wpXml = @"
                    <webParts>
                        <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
                            <metaData><type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" /></metaData>
                        <data>
                          <properties>
                          </properties>
                        </data>
                      </webPart>
                    </webParts>";

            var postProcessor = WebPartPostProcessorFactory.Resolve(wpXml);

            Assert.IsTrue(postProcessor is XsltWebPartPostProcessor);
        }
    }
}
