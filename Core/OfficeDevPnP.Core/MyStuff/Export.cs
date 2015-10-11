using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Enums;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Diagnostics;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that provides generic list creation and manipulation methods
    /// </summary>
    public static partial class ExportExtensions
    {
        public static string Export(this List list)
        {
            var ctx = list.Context;
            var allItems = list.GetItems(CamlQuery.CreateAllItemsQuery());
            ctx.Load(allItems);
            ctx.ExecuteQueryRetry();

            var xml = new XmlDocument();
            var root = xml.CreateElement("pnp", "DataRows", "http://schemas.dev.office.com/PnP/2015/05/ProvisioningSchema");
            xml.AppendChild(root);

            foreach (var item in allItems)
            {
                //item.EnsureProperty(x => x.FieldValues);
                var row = xml.CreateElement("pnp", "DataRow", "http://schemas.dev.office.com/PnP/2015/05/ProvisioningSchema");

                foreach (var fValue in item.FieldValues)
                {
                    var data = xml.CreateElement("pnp", "DataValue", "http://schemas.dev.office.com/PnP/2015/05/ProvisioningSchema");
                    data.SetAttribute("FieldName", fValue.Key);
                    var inner = fValue.Value?.ToString();
                    if (inner != null && inner.Contains("<"))
                    {
                        var cdata = xml.CreateCDataSection(inner);
                        data.AppendChild(cdata);
                    }
                    else
                    {
                        data.InnerXml = inner;
                    }
                    row.AppendChild(data);
                }

                root.AppendChild(row);
            }


            return xml.OuterXml;
        }
    }
}
