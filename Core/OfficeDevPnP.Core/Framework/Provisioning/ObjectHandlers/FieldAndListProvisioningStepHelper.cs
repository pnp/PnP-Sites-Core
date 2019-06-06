using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Field = OfficeDevPnP.Core.Framework.Provisioning.Model.Field;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    public static class FieldAndListProvisioningStepHelper
    {
        static readonly Dictionary<Field,XElement> _fieldXmlDictionary = new Dictionary<Field, XElement>();
        internal static Step GetFieldProvisioningStep(this Field templateField, TokenParser parser)
        {
            XElement schemaElement;
            if (!_fieldXmlDictionary.TryGetValue(templateField, out schemaElement)) {
                schemaElement = XElement.Parse(parser.ParseXmlString(templateField.SchemaXml));
                _fieldXmlDictionary[templateField] = schemaElement;
            }
            var type = (string)schemaElement.Attribute("Type");
            if (type != "Lookup" && type != "LookupMulti")
            {
                return Step.ListAndStandardFields;
            }
            return Step.LookupFields;
        }

        internal static Guid GetFieldId(this Field templateField, TokenParser parser)
        {
            XElement schemaElement;
            if (!_fieldXmlDictionary.TryGetValue(templateField, out schemaElement))
            {
                schemaElement = XElement.Parse(parser.ParseXmlString(templateField.SchemaXml));
                _fieldXmlDictionary[templateField] = schemaElement;
            }
            var id = (Guid)schemaElement.Attribute("ID");
            return id;
        }

        internal static XElement GetSchemaXml(this Field templateField, TokenParser parser, params string[] tokensToSkip)
        {
            XElement schemaElement;
            if (!_fieldXmlDictionary.TryGetValue(templateField, out schemaElement))
            {
                schemaElement = XElement.Parse(parser.ParseXmlString(templateField.SchemaXml, tokensToSkip));
                _fieldXmlDictionary[templateField] = schemaElement;
            }
            return schemaElement;
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