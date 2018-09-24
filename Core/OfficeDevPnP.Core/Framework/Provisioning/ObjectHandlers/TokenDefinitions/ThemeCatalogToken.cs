using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{themecatalog}",
      Description = "Returns the server relative url of the theme catalog",
      Example = "{themecatalog}",
      Returns = "/sites/sitecollection/_catalogs/theme")]
    internal class ThemeCatalogToken : TokenDefinition
    {
        public ThemeCatalogToken(Web web)
            : base(web, "{themecatalog}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Site.EnsureProperty(p => p.RootWeb);
                var catalog = TokenContext.Site.RootWeb.GetCatalog((int)ListTemplateType.ThemeCatalog);
                TokenContext.Load(catalog, c => c.RootFolder.ServerRelativeUrl);
                TokenContext.ExecuteQueryRetry();
                CacheValue = catalog.RootFolder.ServerRelativeUrl;
            }
            return CacheValue;
        }
    }
}