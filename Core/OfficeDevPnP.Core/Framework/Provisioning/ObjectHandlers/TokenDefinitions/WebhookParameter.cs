using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    
    internal class WebhookParameter : SimpleTokenDefinition
    {
        private string _value = null;

        public WebhookParameter(string name, string value)
            : base($"{{webhookparam:{Regex.Escape(name)}}}", $"{{webhookparameter:{Regex.Escape(name)}}}")
        {
            _value = value;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}
