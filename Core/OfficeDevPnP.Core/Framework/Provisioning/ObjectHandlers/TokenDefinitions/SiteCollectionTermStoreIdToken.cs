using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectiontermstoreid}",
        Description = "Returns the id of the given default site collection term store",
        Example = "{sitecollectiontermstoreid}",
        Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
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
                TaxonomySession session = TaxonomySession.GetTaxonomySession(TokenContext);
                var termStore = session.GetDefaultSiteCollectionTermStore();
                TokenContext.Load(termStore, t => t.Id);
                TokenContext.ExecuteQueryRetry();
                if (termStore != null)
                {
                    CacheValue = termStore.Id.ToString();
                }
            }
            return CacheValue;
        }
    }
}