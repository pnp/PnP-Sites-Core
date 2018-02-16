using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
#if !ONPREMISES    
    internal class SiteCollectionConnectedOffice365GroupId : TokenDefinition
    {
        public SiteCollectionConnectedOffice365GroupId(Web web)
            : base(web, "{sitecollectionconnectedoffice365groupid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    context.Load(context.Site, s => s.GroupId);
                    context.ExecuteQueryRetry();
                    if (context.Site.GroupId != null && !context.Site.GroupId.Equals(Guid.Empty))
                    {
                        CacheValue = context.Site.GroupId.ToString();
                    }
                    else
                    {
                        CacheValue = "";
                    }
                }
            }
            return CacheValue;
        }
    }
#endif
}
