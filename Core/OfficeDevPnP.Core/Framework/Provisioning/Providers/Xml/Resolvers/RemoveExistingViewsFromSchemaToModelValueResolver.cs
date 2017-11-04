using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves the RemoveExistingViews attribute from Schema to Domain Model
    /// </summary>
    internal class RemoveExistingViewsFromSchemaToModelValueResolver : IValueResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            var result = false;

            var views = source.GetPublicInstancePropertyValue("Views");
            var removeExistingViews = views?.GetPublicInstancePropertyValue("RemoveExistingViews");

            if (null != removeExistingViews)
            {
                result = (Boolean)removeExistingViews;
            }

            return (result);
        }
    }
}
