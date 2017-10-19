using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Xml.Linq;
using Field = OfficeDevPnP.Core.Framework.Provisioning.Model.Field;

namespace OfficeDevPnP.Core.Extensions
{
    internal static class TemplateFieldExtension
    {
        internal static FieldStage GetFieldStage(this Field templateField, TokenParser parser)
        {
            var schemaElement = XElement.Parse(parser.ParseString(templateField.SchemaXml));
            var type = (string)schemaElement.Attribute("Type");
            var fieldRef = (string)schemaElement.Attribute("FieldRef");
            if (type != "Lookup" && type != "LookupMulti") return FieldStage.Default;

            if (fieldRef != null)
            {
                return FieldStage.DependentLookupFields;
            }
            else
            {
                return FieldStage.LookupFields;
            }
        }
    }
}