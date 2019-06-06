using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectLookupFieldsTests
    {
        //private const string ElementSchema = @"<Field xmlns=""http://schemas.microsoft.com/sharepoint/v3"" StaticName=""DemoField"" DisplayName=""Test Field"" Type=""Text"" ID=""{7E5E53E4-86C2-4A64-9F2E-FDFECE6219E0}"" Group=""PnP"" Required=""true""/>";

        private const string ElementSchema =
            @"<Field Type=""LookupMulti"" DisplayName=""TestLookupField DisplayName"" Required=""FALSE"" EnforceUniqueValues=""FALSE"" List=""Lists/ObjectFieldTestList"" ShowField=""Title"" Mult=""TRUE"" Sortable=""FALSE"" UnlimitedLengthInDocumentLibrary=""FALSE"" Group=""Test Fields"" ID=""{4FE8A6EF-8D50-458A-8870-D473D35E8F69}"" Name=""TestLookupField"" StaticName=""TestLookupField"" />";
        private Guid fieldId = Guid.Parse("{4FE8A6EF-8D50-458A-8870-D473D35E8F69}");
        private const string ListTitle = "ObjectFieldTestList";
        private string _listIdWithBraces = string.Empty;

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                var f = ctx.Web.GetFieldById<FieldLookup>(fieldId); // Guid matches ID in field caml.
                if (f != null)
                {
                    f.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
                var list = ctx.Web.GetListByTitle(ListTitle);
                if (list != null)
                {
                    list.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
            }
        }

        [TestInitialize]
        public void Initialize()
        {
            CleanUp(); // just in case ..
            using (var ctx = TestCommon.CreateClientContext())
            {
                // create our list that we do lookups against
                var sourceList = ctx.Web.CreateList(ListTemplateType.GenericList, listName: ListTitle,
                    enableVersioning: false, urlPath: "Lists/" + ListTitle);
                ctx.Load(sourceList, x=>x.Id);
                ctx.ExecuteQueryRetry();
                // listId with braces... this is how SharePoint maintains the list reference in the schema xml in the lookup field
                _listIdWithBraces = sourceList.Id.ToString("B"); 


            }
        }

        //TODO: redesign this test case after refactoring work done for lookup fields
        //[TestMethod]
        //public void CanProvisionObjects()
        //{
        //    var template = new ProvisioningTemplate();
        //    template.SiteFields.Add(new Core.Framework.Provisioning.Model.Field() { SchemaXml = ElementSchema });

        //    using (var ctx = TestCommon.CreateClientContext())
        //    {
        //        var parser = new TokenParser(ctx.Web, template);
        //        new ObjectField(FieldAndListProvisioningStepHelper.Step.LookupFields).ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());
        //        new ObjectLookupFields().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

        //        var f = ctx.Web.GetFieldById<FieldLookup>(fieldId);

        //        Assert.IsNotNull(f);
        //        Assert.IsInstanceOfType(f, typeof(FieldLookup));

        //        var schemaXml = f.SchemaXml;
        //        // so listId MUST have braces
        //        Assert.IsTrue(schemaXml.Contains("List=\""+_listIdWithBraces+"\""));
        //        // web id should NOT have braces
        //        Assert.IsTrue(schemaXml.Contains("WebId=\"" + ctx.Web.Id.ToString()+ "\""));
        //        // Source ID MUST have braces
        //        Assert.IsTrue(schemaXml.Contains("SourceID=\"" + ctx.Web.Id.ToString("B") + "\""));
        //    }

        //}
    }
}
