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
    /// Type resolver for Site Footer Link from model to schema
    /// </summary>
    internal class SiteFooterLinkFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Array result = null;

            Object modelSource = source as Model.SiteFooter;
            if (modelSource == null)
            {
                modelSource = source as Model.SiteFooterLink;
            }
            
            if (modelSource != null)
            {
                Model.SiteFooterLinkCollection sourceLinks = modelSource.GetPublicInstancePropertyValue("FooterLinks") as Model.SiteFooterLinkCollection;
                if (sourceLinks != null && sourceLinks.Count > 0)
                {
                    var siteFooterLinkTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.FooterLink, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var siteFooterLinkType = Type.GetType(siteFooterLinkTypeName, true);

                    result = Array.CreateInstance(siteFooterLinkType, sourceLinks.Count);

                    resolvers = new Dictionary<string, IResolver>();
                    resolvers.Add($"{siteFooterLinkType}.FooterLink1", new SiteFooterLinkFromModelToSchemaTypeResolver());

                    for (Int32 c = 0; c < sourceLinks.Count; c++)
                    {
                        var targetFooterLinkItem = Activator.CreateInstance(siteFooterLinkType);
                        PnPObjectsMapper.MapProperties(sourceLinks[c], targetFooterLinkItem, resolvers, recursive);

                        result.SetValue(targetFooterLinkItem, c);
                    }
                }
            }

            return (result);
        }
    }
}
