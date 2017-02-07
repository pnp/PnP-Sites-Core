using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    public static class SerializerExtensions
    {
        public static bool GetOptionalBoolValue(this XElement element, XName attributeName, bool defaultValue = false)
        {
            var returnValue = defaultValue;
            var attribute = element.Attribute(attributeName);
            if (attribute != null)
            {
                bool.TryParse(attribute.Value, out returnValue);
            }
            return returnValue;
        }

        public static void AddOptionalAttribute(this XElement element, string attributeName, string value)
        {
            if (value != null)
            {
                var attribute = new XAttribute(attributeName, value);
                element.Add(attribute);
            }
        }

        public static void AddOptionalAttribute(this XElement element, string attributeName, bool value)
        {
            var attribute = new XAttribute(attributeName, value);
            element.Add(attribute);
        }
    }
}
