using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionTermStoreIdToken : TokenDefinition
    {
        public SiteCollectionTermStoreIdToken(Web web)
            : base(web, "~sitecollectiontermstoreid", "{sitecollectiontermstoreid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    TaxonomySession session = TaxonomySession.GetTaxonomySession(context);
                    var termStore = session.GetDefaultSiteCollectionTermStore();
                    context.Load(termStore, t => t.Id);
                    context.ExecuteQueryRetry();
                    if (termStore != null)
                    {
                        CacheValue = termStore.Id.ToString();
                    }
                }
            }
            return CacheValue;
        }
    }
}