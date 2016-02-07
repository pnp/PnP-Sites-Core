using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{
#if !CLIENTSDKV15
    internal static class UserResourceExtensions
    {
        public static bool SetUserResourceValue(this UserResource userResource, string tokenValue, TokenParser parser)
        {
            bool isDirty = false;

            if (userResource != null && !String.IsNullOrEmpty(tokenValue))
            {
                var resourceValues = parser.GetResourceTokenResourceValues(tokenValue);
                foreach (var resourceValue in resourceValues)
                {
                    userResource.SetValueForUICulture(resourceValue.Item1, resourceValue.Item2);
                    isDirty = true;
                }
            }

            return isDirty;
        }

        public static bool ContainsResourceToken(this string value)
        {
            if (value != null)
            {
                value = value.ToLower();
                return value.IndexOf("{res:") > -1 ||
                    value.IndexOf("{loc:") > -1 ||
                    value.IndexOf("{resource:") > -1 ||
                    value.IndexOf("{localize:") > -1 ||
                    value.IndexOf("{localization:") > -1;
            }
            else
            {
                return (false);
            }
        }
    }
#endif
}
