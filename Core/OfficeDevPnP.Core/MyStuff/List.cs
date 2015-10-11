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
    public static partial class ListExtensions
    {
        public static void ReIndexList(this List list)
        {
            int searchversion = 0;
            if (list.PropertyBagContainsKey("vti_searchversion"))
            {
                searchversion = (int)list.GetPropertyBagValueInt("vti_searchversion", 0);
            }
            list.SetPropertyBagValue("vti_searchversion", searchversion + 1);
        }
    }
}
