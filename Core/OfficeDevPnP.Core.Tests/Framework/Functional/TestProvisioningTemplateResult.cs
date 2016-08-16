using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    public class TestProvisioningTemplateResult
    {
        public ProvisioningTemplate SourceTemplate { get; set; }
        public TokenParser SourceTokenParser { get; set; }
        public ProvisioningTemplate TargetTemplate { get; set; }
        public TokenParser TargetTokenParser { get; set; }

    }
}
