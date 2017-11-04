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
    /// Resolves a list of Views from Schema to Domain Model
    /// </summary>
    internal class ListViewsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public ListViewsFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new List<Model.View>();

            var views = source.GetPublicInstancePropertyValue("Views");
            var xmlAny = views?.GetPublicInstancePropertyValue("Any") as XmlElement[];

            if (null != xmlAny)
            {
                result.AddRange(
                    from x in xmlAny select 
                    new Model.View { SchemaXml = x.OuterXml });
            }

            return (result);
        }
    }
}
