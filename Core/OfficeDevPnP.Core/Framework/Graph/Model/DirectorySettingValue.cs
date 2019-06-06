using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Represents a single value of a Directory Setting
    /// </summary>
    public class DirectorySettingValue
    {
        public String DefaultValue { get; set; }

        public String Description { get; set; }

        public String Name { get; set; }

        public String Type { get; set; }

        public String Value { get; set; }
    }
}
