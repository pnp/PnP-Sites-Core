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
            : base(web, "~themecatalog", "{themecatalog}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                using (ClientContext cc = Web.Context.GetSiteCollectionContext())
                {
                    var catalog = cc.Web.GetCatalog((int)ListTemplateType.ThemeCatalog);
                    cc.Load(catalog, c => c.RootFolder.ServerRelativeUrl);
                    cc.ExecuteQueryRetry();
                    CacheValue = catalog.RootFolder.ServerRelativeUrl;
                }
            }
            return CacheValue;
        }
    }
}