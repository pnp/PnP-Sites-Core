using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    internal class ODataBindJsonResolver : CamelCasePropertyNamesContractResolver
    {
        protected override JsonProperty CreateProperty(MemberInfo member, MemberSerialization memberSerialization)
        {
            var property = base.CreateProperty(member, memberSerialization);
            var skipEmptyArray = false;
            var originalPropertyName = property.PropertyName;

            if (property.PropertyName.Equals("template_odata_bind", StringComparison.OrdinalIgnoreCase))
            {
                property.PropertyName = "template@odata.bind";
            }
            else if (property.PropertyName.Equals("owners_odata_bind", StringComparison.OrdinalIgnoreCase))
            {
                property.PropertyName = "owners@odata.bind";
                skipEmptyArray = true;
            }
            else if (property.PropertyName.Equals("members_odata_bind", StringComparison.OrdinalIgnoreCase))
            {
                property.PropertyName = "members@odata.bind";
                skipEmptyArray = true;
            }
            else if (property.PropertyName.Equals("private_channel_member_odata_type", StringComparison.OrdinalIgnoreCase))
            {
                property.PropertyName = "@odata.type";
            }
            else if (property.PropertyName.Equals("private_channel_user_odata_bind", StringComparison.OrdinalIgnoreCase))
            {
                property.PropertyName = "user@odata.bind";
            }

            if (skipEmptyArray)
            {
                property.ShouldSerialize = instance =>
                {
                    var enumerator = instance
                        .GetType()
                        .GetProperty(originalPropertyName)
                        .GetValue(instance, null) as IEnumerable;

                    if (enumerator != null)
                    {
                        // check to see if there is at least one item in the Enumerable
                        return enumerator.GetEnumerator().MoveNext();
                    }
                    else
                    {
                        // if the enumerable is null, we defer the decision to NullValueHandling
                        return true;
                    }
                };
            }

            return property;
        }
    }
}
