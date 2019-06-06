using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class FieldImplementation : ImplementationBase
    {
        internal void SiteCollectionFieldAdding(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Ensure we can test clean
                DeleteFields(cc);

                // Add fields
                var result = TestProvisioningTemplate(cc, "field_add.xml", Handlers.Fields | Handlers.TermGroups);
                FieldValidator pv = new FieldValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.SiteFields, result.TargetTemplate.SiteFields, result.TargetTokenParser));

                // Apply delta to fields
                var result2 = TestProvisioningTemplate(cc, "field_delta_1.xml", Handlers.Fields);
                Assert.IsTrue(pv.Validate(result2.SourceTemplate.SiteFields, result2.TargetTemplate.SiteFields, result2.TargetTokenParser));
            }
        }

        #region Helper methods
        private void DeleteFields(ClientContext cc)
        {
            cc.Load(cc.Web.Fields, f => f.Include(t => t.InternalName));
            cc.ExecuteQueryRetry();

            foreach (var field in cc.Web.Fields.ToList())
            {
                // First drop the fields that have 2 _'s...convention used to name the fields dependent on a lookup.
                if (field.InternalName.Replace("FLD_", "").IndexOf("_") > 0)
                {
                    if (field.InternalName.StartsWith("FLD_"))
                    {
                        field.DeleteObject();
                    }
                }
            }

            foreach (var field in cc.Web.Fields.ToList())
            {
                if (field.InternalName.StartsWith("FLD_"))
                {
                    field.DeleteObject();
                }
            }

            cc.ExecuteQueryRetry();

        }
        #endregion
    }
}