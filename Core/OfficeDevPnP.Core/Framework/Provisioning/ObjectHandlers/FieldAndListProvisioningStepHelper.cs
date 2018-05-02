using System.Xml.Linq;
using Field = OfficeDevPnP.Core.Framework.Provisioning.Model.Field;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    public static class FieldAndListProvisioningStepHelper
    {
        internal static Step GetFieldProvisioningStep(this Field templateField, TokenParser parser)
        {
            var schemaElement = XElement.Parse(parser.ParseString(templateField.SchemaXml));
            var type = (string)schemaElement.Attribute("Type");
            if (type != "Lookup" && type != "LookupMulti")
            {
                return Step.ListAndStandardFields;
            }
            else
            {
                return Step.LookupFields;
            }
        }

        public enum Step
        {
            /// <summary>
            /// The list itself and fields that aren't lookup fields are provisioned
            /// </summary>
            ListAndStandardFields,

            /// <summary>
            /// Focus on lookup fields. This assumes target lists are yet available
            /// </summary>
            LookupFields,

            /// <summary>
            /// Remaining list customization
            /// </summary>
            ListSettings,
            /// <summary>
            /// The handler is exporting
            /// </summary>
            Export
        }
    }
}