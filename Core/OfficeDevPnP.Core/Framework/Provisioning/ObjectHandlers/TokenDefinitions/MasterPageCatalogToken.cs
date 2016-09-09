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
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    List catalog;
                    // Check if the current web is a sub-site
                    if (Web.IsSubSite())
                    {
                        // Master page URL needs to be retrieved from the rootweb
                        var rootWeb = context.Site.RootWeb;
                        catalog = rootWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                    }
                    else
                    {
                        catalog = context.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                    }

                    context.Load(catalog, c => c.RootFolder.ServerRelativeUrl);
                    context.ExecuteQueryRetry();
                    CacheValue = catalog.RootFolder.ServerRelativeUrl;
                }
            }
            return CacheValue;
        }
    }
}