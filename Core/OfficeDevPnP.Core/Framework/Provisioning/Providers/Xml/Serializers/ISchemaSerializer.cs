using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    internal interface ISchemaSerializer
    {
        XElement FromProvisioningTemplate(ProvisioningTemplate template, XNamespace ns);

        ProvisioningTemplate ToProvisioningTemplate(XElement templateElement, XNamespace ns, ProvisioningTemplate template);
    }
}
