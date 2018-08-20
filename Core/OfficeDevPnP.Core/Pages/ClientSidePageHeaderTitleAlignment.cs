using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES || SP2019
    public enum ClientSidePageHeaderTitleAlignment
    {
        Center,
        Left
    }
#endif
}
