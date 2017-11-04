using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class ClientSidePagesValidator : ValidatorBase
    {
        public ClientSidePagesValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05;
        }

        public bool Validate(ClientSidePageCollection sourcePages, Microsoft.SharePoint.Client.ClientContext ctx)
        {
            return true;
        }
    }
}
