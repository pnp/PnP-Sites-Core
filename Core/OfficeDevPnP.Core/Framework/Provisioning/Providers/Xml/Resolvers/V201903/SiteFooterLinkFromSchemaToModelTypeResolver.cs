using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for Site Footer Link from schema to model
    /// </summary>
    internal class SiteFooterLinkFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.SiteFooterLink>();

            var links = source.GetPublicInstancePropertyValue("FooterLinks");
            if (null == links)
            {
                links = source.GetPublicInstancePropertyValue("FooterLink1");
            }

            resolvers = new Dictionary<string, IResolver>();
            resolvers.Add($"{typeof(Model.SiteFooterLink).FullName}.FooterLinks", new SiteFooterLinkFromSchemaToModelTypeResolver());

            if (null != links)
            {
                foreach (var f in ((IEnumerable)links))
                {
                    var targetItem = new Model.SiteFooterLink();
                    PnPObjectsMapper.MapProperties(f, targetItem, resolvers, recursive);
                    result.Add(targetItem);
                }
            }

            return (result);
        }
    }
}
