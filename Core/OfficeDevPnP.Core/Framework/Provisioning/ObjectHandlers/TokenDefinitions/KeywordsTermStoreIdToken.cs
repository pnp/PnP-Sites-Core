using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class KeywordsTermStoreIdToken : TokenDefinition
    {
        public KeywordsTermStoreIdToken(Web web)
            : base(web, "~keywordstermstoreid", "{keywordstermstoreid}")
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
                    var termStore = session.GetDefaultKeywordsTermStore();
                    context.Load(termStore, t => t.Id);
                    context.ExecuteQueryRetry();
                    CacheValue = termStore.Id.ToString();
                }
            }
            return CacheValue;
        }
    }
}