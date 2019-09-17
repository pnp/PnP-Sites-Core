using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Drive
{
    /// <summary>
    /// Defines a DriveItem object
    /// </summary>
    public partial class DriveItem : DriveItemBase
    {
        protected override bool EqualsInherited(DriveItemBase other)
        {
            if (!(other is DriveItem otherTyped))
            {
                return (false);
            }

            // At the moment we don't have anything to compare
            return (true);
        }

        protected override int GetInheritedHashCode()
        {
            // At the moment we don't have an hashcode for this specialized type
            return ((String.Empty).GetHashCode());
        }
    }
}
