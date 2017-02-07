using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    public class SupportedTemplateSchemasAttribute : Attribute
    {
        public SupportedSchema Schemas { get; set; }
        public int Order { get; set; }
    }

    [Flags]
    public enum SupportedSchema
    {
        V201605,
        V201703
    }
}
