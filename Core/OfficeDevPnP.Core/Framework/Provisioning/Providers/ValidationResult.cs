using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers
{
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<Exception> Exceptions { get; set; }
    }
}
