using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class MasterPageCatalogToken : TokenDefinition
    {
        public MasterPageCatalogToken(Web web)
            : base(web, "~masterpagecatalog","{masterpagecatalog}")
        {
        }

        public override string GetReplaceValue()
        {
            if (this.CacheValue == null)
            {
                List catalog;
                // Check if the current web is a sub-site
                if (Web.IsSubSite())
                {
                    // Master page URL needs to be retrieved from the rootweb
                    var rootWeb = (Web.Context as ClientContext).Site.RootWeb;
                    catalog = rootWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                }
                else
                {
                    catalog = Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                }
                Web.Context.Load(catalog, c => c.RootFolder.ServerRelativeUrl);
                Web.Context.ExecuteQueryRetry();
                CacheValue = catalog.RootFolder.ServerRelativeUrl;
            }
            return CacheValue;
        }
    }
}