using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{keywordstermstoreid}",
     Description = "Returns a id of the default keywords term store",
     Example = "{keywordstermstoreid}",
     Returns = "f2cd6d5b-1391-480e-a3dc-7f7f96137382")]
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
                TaxonomySession session = TaxonomySession.GetTaxonomySession(TokenContext);
                var termStore = session.GetDefaultKeywordsTermStore();
                TokenContext.Load(termStore, t => t.Id);
                TokenContext.ExecuteQueryRetry();
                CacheValue = termStore.Id.ToString();
            }
            return CacheValue;
        }
    }
}