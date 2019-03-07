using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    public enum ClientSidePageHeaderLayoutType
    {
        FullWidthImage,
        NoImage,
#if !SP2019
        ColorBlock,
        CutInShape,
#endif
    }
#endif
    }
