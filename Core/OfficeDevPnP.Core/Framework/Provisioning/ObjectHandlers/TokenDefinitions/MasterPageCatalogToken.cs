using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{masterpagecatalog}",
     Description = "Returns a server relative url of the master page catalog",
     Example = "{masterpagecatalog}",
     Returns = "/sites/mysite/_catalogs/masterpage")]
    internal class MasterPageCatalogToken : TokenDefinition
    {
        public MasterPageCatalogToken(Web web)
            : base(web, "~masterpagecatalog", "{masterpagecatalog}")
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
                    var rootWeb = TokenContext.Site.RootWeb;
                    catalog = rootWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                }
                else
                {
                    catalog = TokenContext.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                }

                TokenContext.Load(catalog, c => c.RootFolder.ServerRelativeUrl);
                TokenContext.ExecuteQueryRetry();
                CacheValue = catalog.RootFolder.ServerRelativeUrl;
            }
            return CacheValue;
        }
    }
}