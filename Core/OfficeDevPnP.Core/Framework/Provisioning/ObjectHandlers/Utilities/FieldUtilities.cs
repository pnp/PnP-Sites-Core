using Microsoft.SharePoint.Client;
using System;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    public static class FieldUtilities
    {
        public static string FixLookupField(string fieldXml, Web web)
        {
            var fieldElement = XElement.Parse(fieldXml);
            if ((string)fieldElement.Attribute("Type") == "Lookup")
            {
                var listAttr = (string)fieldElement.Attribute("List");
                if (!Guid.TryParse(listAttr, out Guid _))
                {
                    var targetList = web.GetList($"{web.ServerRelativeUrl.TrimEnd('/')}/{listAttr}");
                    fieldElement.SetAttributeValue("List", targetList.EnsureProperty(l => l.Id).ToString("B"));
                    return fieldElement.ToString();
                }
            }

            return fieldXml;
        }
    }
}