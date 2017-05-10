using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Decimal value into a Double
    /// </summary>
    internal class FromBasePermissionsToStringValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        public object Resolve(object source, object destination, object sourceValue)
        {
            string res = null;
            if (sourceValue != null)
            {
                var basePermissions = (BasePermissions)sourceValue;
                List<string> permissions = new List<string>();
                foreach (var pk in (PermissionKind[])Enum.GetValues(typeof(PermissionKind)))
                {
                    if (basePermissions.Has(pk) && pk != PermissionKind.EmptyMask)
                    {
                        permissions.Add(pk.ToString());
                    }
                }
                res = string.Join(",", permissions.ToArray());
            }
            return res;
        }
    }
}
