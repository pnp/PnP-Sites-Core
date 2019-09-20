using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201903
{
    /// <summary>
    /// Resolves SiteDesign.WebTemplate enum to compatible type
    /// </summary>
    internal class TenantSiteDesignsWebTemplateFromModelToSchemaValueResolver : IValueResolver
    {
        public string Name => GetType().Name;

        public object Resolve(object source, object destination, object sourceValue)
        {
            if (sourceValue is SiteDesignWebTemplate && (SiteDesignWebTemplate)sourceValue == SiteDesignWebTemplate.TeamSite)
            {
                return 0;
            }

            if (sourceValue is SiteDesignWebTemplate && (SiteDesignWebTemplate)sourceValue == SiteDesignWebTemplate.CommunicationSite)
            {
                return 1;
            }

            return sourceValue;
        }
    }
}