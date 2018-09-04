using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Attributes
{
    [AttributeUsage(AttributeTargets.Class,
                       AllowMultiple = true)]
    public sealed class TokenDefinitionDescriptionAttribute : Attribute
    {
        public string Token { get; set; }
        public string Description { get; set; }
        public string Returns { get; set; }
        public string Example { get; set; }
    }
}
