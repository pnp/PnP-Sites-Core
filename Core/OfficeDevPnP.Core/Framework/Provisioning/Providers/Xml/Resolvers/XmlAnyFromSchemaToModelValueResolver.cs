using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Dictionary into an Array of objects
    /// </summary>
    internal class XmlAnyFromSchemaToModelValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        private String _elementName;

        public XmlAnyFromSchemaToModelValueResolver(String elementName)
        {
            this._elementName = elementName;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            XElement result = null;

            var xmlAny = sourceValue.GetPublicInstancePropertyValue("Any") as XmlElement[];

            if (null != xmlAny)
            {
                result = new XElement(this._elementName,
                    from x in xmlAny select x.ToXElement());
            }

            return (result);
        }
    }
}
