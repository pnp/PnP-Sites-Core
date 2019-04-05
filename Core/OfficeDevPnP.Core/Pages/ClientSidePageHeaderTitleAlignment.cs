using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    public enum ClientSidePageHeaderTitleAlignment
    {
        Center,
        Left
    }
#endif
}
