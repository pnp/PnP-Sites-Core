using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionTermGroupNameToken : TokenDefinition
    {
        public SiteCollectionTermGroupNameToken(Web web)
            : base(web, "~sitecollectiontermgroupname", "{sitecollectiontermgroupname}")
        {
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                // The token is requested. Check if the group exists and if not, create it
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    var site = context.Site;
                    var session = TaxonomySession.GetTaxonomySession(context);
                    var termstore = session.GetDefaultSiteCollectionTermStore();
                    var termGroup = termstore.GetSiteCollectionGroup(site, true);
                    context.Load(termGroup);
                    context.ExecuteQueryRetry();

                    CacheValue = termGroup.Name.ToString();
                }
            }
            return CacheValue;
        }
    }
}