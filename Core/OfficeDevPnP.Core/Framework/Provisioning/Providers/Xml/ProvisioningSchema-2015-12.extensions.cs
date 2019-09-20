using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201512
{
    [XmlSchemaProviderAttribute("GetSchema")]
    public partial class WikiPageWebPart : IXmlSerializable
    {
        public static XmlQualifiedName GetSchema(XmlSchemaSet schemaSet)
        {
            String wikiPageWebPartSchemaString = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
              "<xsd:schema targetNamespace=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "elementFormDefault=\"qualified\" " +
                "xmlns=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "xmlns:pnp=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" +
                    "<xsd:complexType name=\"WikiPageWebPart\">" +
                        "<xsd:all>" +
                            "<xsd:element name=\"Contents\" minOccurs=\"1\" maxOccurs=\"1\">" +
                                "<xsd:complexType>" +
                                    "<xsd:sequence>" +
                                        "<xsd:any processContents=\"lax\" namespace=\"##any\" minOccurs=\"0\" />" +
                                    "</xsd:sequence>" +
                                "</xsd:complexType>" +
                            "</xsd:element>" +
                        "</xsd:all>" +
                        "<xsd:attribute name=\"Title\" type=\"xsd:string\" use=\"required\" />" +
                        "<xsd:attribute name=\"Row\" type=\"xsd:int\" use=\"required\" />" +
                        "<xsd:attribute name=\"Column\" type=\"xsd:int\" use=\"required\" />" +
                    "</xsd:complexType>" +
                "</xsd:schema>";

            XmlSchema webPartSchema = XmlSchema.Read(new StringReader(wikiPageWebPartSchemaString), null);
            schemaSet.XmlResolver = new XmlUrlResolver();
            schemaSet.Add(webPartSchema);

            return (new XmlQualifiedName("WikiPageWebPart",
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                ));
        }

        XmlSchema IXmlSerializable.GetSchema()
        {
            throw new NotImplementedException("This method should never be called. We implemented the static GetSchema method for XmlSchemaProviderAttribute.");
        }

        void IXmlSerializable.ReadXml(XmlReader reader)
        {
            XNamespace ns =
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                ;

            XElement webPartXml = (XElement)XElement.ReadFrom(reader);
            this.Title = webPartXml.Attribute("Title").Value;
            this.Row = Int32.Parse(webPartXml.Attribute("Row").Value);
            this.Column = Int32.Parse(webPartXml.Attribute("Column").Value);

            XElement webPartContents = webPartXml.Element(ns + "Contents");
            this.Contents = webPartContents.ToXmlElement();
        }

        void IXmlSerializable.WriteXml(XmlWriter writer)
        {
            writer.WriteAttributeString("Title", this.Title);
            writer.WriteAttributeString("Row", this.Row.ToString());
            writer.WriteAttributeString("Column", this.Column.ToString());
            writer.WriteStartElement(XMLConstants.PROVISIONING_SCHEMA_PREFIX, "Contents",
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                );

            using (XmlReader xr = new XmlNodeReader(this.Contents))
            {
                writer.WriteNode(xr, false);
            }

            writer.WriteEndElement();
        }
    }

    [XmlSchemaProviderAttribute("GetSchema")]
    public partial class BaseFieldValue : IXmlSerializable
    {
        public static XmlQualifiedName GetSchema(XmlSchemaSet schemaSet)
        {
            String baseFieldValueSchemaString = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
              "<xsd:schema targetNamespace=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "elementFormDefault=\"qualified\" " +
                "xmlns=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "xmlns:pnp=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" +
                    "<xsd:complexType name=\"BaseFieldValue\">" +
                        "<xsd:simpleContent>" +
                            "<xsd:extension base=\"xsd:string\">" +
                                "<xsd:attribute name=\"FieldName\" use=\"required\" type=\"xsd:string\"/>" +
                            "</xsd:extension>" +
                        "</xsd:simpleContent>" +
                    "</xsd:complexType>" +
                "</xsd:schema>";

            XmlSchema baseFieldValueSchema = XmlSchema.Read(new StringReader(baseFieldValueSchemaString), null);
            schemaSet.XmlResolver = new XmlUrlResolver();
            schemaSet.Add(baseFieldValueSchema);

            return (new XmlQualifiedName("BaseFieldValue",
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                ));
        }

        XmlSchema IXmlSerializable.GetSchema()
        {
            throw new NotImplementedException("This method should never be called. We implemented the static GetSchema method for XmlSchemaProviderAttribute.");
        }

        void IXmlSerializable.ReadXml(XmlReader reader)
        {
            XNamespace ns =
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                ;

            XElement baseFieldValueXml = (XElement)XElement.ReadFrom(reader);
            this.FieldName = baseFieldValueXml.Attribute("FieldName").Value;
            this.Value = baseFieldValueXml.Value;
        }

        void IXmlSerializable.WriteXml(XmlWriter writer)
        {
            Regex regExHTML = new Regex(@"<(\w|-|_)+>(.)*<\/(\w)+>");

            writer.WriteAttributeString("FieldName", this.FieldName);

            // If the content is HTML-like, use a CDATA section
            if (!String.IsNullOrEmpty(this.Value))
            {
                if (regExHTML.IsMatch(this.Value))
                {
                    writer.WriteCData(this.Value);
                }
                else
                {
                    writer.WriteString(this.Value);
                }
            }
            else
            {
                writer.WriteString(String.Empty);
            }
        }
    }

    [XmlSchemaProviderAttribute("GetSchema")]
    public partial class WebPartPageWebPart : IXmlSerializable
    {
        public static XmlQualifiedName GetSchema(XmlSchemaSet schemaSet)
        {
            String wikiPageWebPartSchemaString = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
              "<xsd:schema targetNamespace=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "elementFormDefault=\"qualified\" " +
                "xmlns=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "xmlns:pnp=\"http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema\" " +
                "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" +
                    "<xsd:complexType name=\"WebPartPageWebPart\">" +
                        "<xsd:all>" +
                            "<xsd:element name=\"Contents\" minOccurs=\"1\" maxOccurs=\"1\">" +
                                "<xsd:complexType>" +
                                    "<xsd:sequence>" +
                                        "<xsd:any processContents=\"lax\" namespace=\"##any\" minOccurs=\"0\" />" +
                                    "</xsd:sequence>" +
                                "</xsd:complexType>" +
                            "</xsd:element>" +
                        "</xsd:all>" +
                        "<xsd:attribute name=\"Title\" type=\"xsd:string\" use=\"required\" />" +
                        "<xsd:attribute name=\"Zone\" type=\"xsd:string\" use=\"required\" />" +
                        "<xsd:attribute name=\"Order\" type=\"xsd:int\" use=\"required\" />" +
                    "</xsd:complexType>" +
                "</xsd:schema>";

            XmlSchema webPartSchema = XmlSchema.Read(new StringReader(wikiPageWebPartSchemaString), null);
            schemaSet.XmlResolver = new XmlUrlResolver();
            schemaSet.Add(webPartSchema);

            return (new XmlQualifiedName("WebPartPageWebPart",
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                ));
        }

        XmlSchema IXmlSerializable.GetSchema()
        {
            throw new NotImplementedException("This method should never be called. We implemented the static GetSchema method for XmlSchemaProviderAttribute.");
        }

        void IXmlSerializable.ReadXml(XmlReader reader)
        {
            XNamespace ns =
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                ;

            XElement webPartXml = (XElement)XElement.ReadFrom(reader);
            this.Title = webPartXml.Attribute("Title").Value;
            this.Zone = webPartXml.Attribute("Zone").Value;
            this.Order = Int32.Parse(webPartXml.Attribute("Order").Value);

            XElement webPartContents = webPartXml.Element(ns + "Contents");
            this.Contents = webPartContents.ToXmlElement();
        }

        void IXmlSerializable.WriteXml(XmlWriter writer)
        {
            writer.WriteAttributeString("Title", this.Title);
            writer.WriteAttributeString("Zone", this.Zone);
            writer.WriteAttributeString("Order", this.Order.ToString());
            writer.WriteStartElement(XMLConstants.PROVISIONING_SCHEMA_PREFIX, 
                "Contents",
#pragma warning disable 0618
                XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12
#pragma warning restore 0618
                );

            using (XmlReader xr = new XmlNodeReader(this.Contents))
            {
                writer.WriteNode(xr, false);
            }

            writer.WriteEndElement();
        }
    }
}
