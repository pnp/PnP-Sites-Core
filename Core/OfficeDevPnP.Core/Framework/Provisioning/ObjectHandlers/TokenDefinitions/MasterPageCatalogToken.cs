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
                    var catalog = context.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                    context.Load(catalog, c => c.RootFolder.ServerRelativeUrl);
                    context.ExecuteQueryRetry();
                    CacheValue = catalog.RootFolder.ServerRelativeUrl;
                }
            }
            return CacheValue;
        }
    }
}