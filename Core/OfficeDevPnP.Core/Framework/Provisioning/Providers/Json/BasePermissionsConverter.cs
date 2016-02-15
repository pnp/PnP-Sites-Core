using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Json
{
    internal class BasePermissionsConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            return (typeof(Microsoft.SharePoint.Client.BasePermissions).IsAssignableFrom(objectType));
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            Microsoft.SharePoint.Client.BasePermissions result =
                new Microsoft.SharePoint.Client.BasePermissions();

            JToken token = JToken.Load(reader);
            String basePermissionString = token.ToString();

            if (!String.IsNullOrEmpty(basePermissionString))
            {
                // Is it an int value (for backwards compability)?
                Int32 permissionInt = 0;
                if (int.TryParse(basePermissionString, out permissionInt))
                {
                    result.Set((Microsoft.SharePoint.Client.PermissionKind)permissionInt);
                }
                else
                {
                    foreach (var pk in basePermissionString.Split(new char[] { ',' }))
                    {
                        Microsoft.SharePoint.Client.PermissionKind permissionKind =
                            Microsoft.SharePoint.Client.PermissionKind.AddAndCustomizePages;
                        if (Enum.TryParse(basePermissionString, out permissionKind))
                        {
                            result.Set(permissionKind);
                        }
                    }
                }
            }

            return (result);
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            String jsonValue = null;

            Microsoft.SharePoint.Client.BasePermissions basePermissions = 
                value as Microsoft.SharePoint.Client.BasePermissions;
            if (basePermissions != null)
            {
                List<String> permissions = new List<String>();
                foreach (var pk in (Microsoft.SharePoint.Client.PermissionKind[])Enum.GetValues(typeof(Microsoft.SharePoint.Client.PermissionKind)))
                {
                    if (basePermissions.Has(pk) && pk !=
                        Microsoft.SharePoint.Client.PermissionKind.EmptyMask)
                    {
                        permissions.Add(pk.ToString());
                    }
                }
                jsonValue = string.Join(",", permissions.ToArray());
            }

            writer.WriteValue(jsonValue);
        }
    }
}
