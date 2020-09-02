using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201909
{
    /// <summary>
    /// Type resolver for Office365GroupsSettings from Schema to Model
    /// </summary>
    internal class Office365GroupsSettingsFromSchemaToModel : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Model.Office365Groups.Office365GroupsSettings result = null;
            var settings = source.GetPublicInstancePropertyValue("Office365GroupsSettings");

            if (null != settings)
            {
                result = new Model.Office365Groups.Office365GroupsSettings();
                foreach (var p in ((IEnumerable)settings))
                {
                    result.Properties.Add(
                        (String)p.GetPublicInstancePropertyValue("Key"),
                        (String)p.GetPublicInstancePropertyValue("Value"));
                }
            }

            return (result);
        }
    }
}
