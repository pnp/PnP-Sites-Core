using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Attributes;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectiontermgroupid}",
        Description = "Returns the id of the site collection term group",
        Example = "{sitecollectiontermgroupid}",
        Returns = "767bc144-e605-4d8c-885a-3a980feb39c6")]
    internal class SiteCollectionTermGroupIdToken : VolatileTokenDefinition
    {
        public SiteCollectionTermGroupIdToken(Web web)
            : base(web, "{sitecollectiontermgroupid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                try
                {
                    // The token is requested. Check if the group exists and if not, create it
                    var site = TokenContext.Site;
                    var session = TaxonomySession.GetTaxonomySession(TokenContext);
                    var termstore = session.GetDefaultSiteCollectionTermStore();
                    var termGroup = termstore.GetSiteCollectionGroup(site, true);
                    TokenContext.Load(termGroup);
                    TokenContext.ExecuteQueryRetry();

                    CacheValue = termGroup.Id.ToString();
                }
                catch(ServerUnauthorizedAccessException)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.TermGroup_No_Access);
                }
            }
            return CacheValue;
        }
    }
}