using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Attributes;
using System;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectiontermgroupname}",
        Description = "Returns the name of the site collection term group",
        Example = "{sitecollectiontermgroupname}",
        Returns = "Site Collection - mytenant.sharepoint.com-sites-mysite")]
    internal class SiteCollectionTermGroupNameToken : VolatileTokenDefinition
    {
        public SiteCollectionTermGroupNameToken(Web web)
            : base(web, "{sitecollectiontermgroupname}")
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

                    CacheValue = termGroup.Name.ToString();
                }
                catch (ServerUnauthorizedAccessException)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.TermGroup_No_Access);
                }
            }
            return CacheValue;
        }
    }
}