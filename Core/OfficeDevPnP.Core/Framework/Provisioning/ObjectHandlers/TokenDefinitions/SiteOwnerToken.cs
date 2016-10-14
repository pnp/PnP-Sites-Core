using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
	internal class SiteOwnerToken : TokenDefinition
	{
		public SiteOwnerToken(Web web)
			: base(web, "~siteowner", "{siteowner}")
		{
		}

		public override string GetReplaceValue()
		{
			if (CacheValue == null)
			{
				var context = Web.Context as ClientContext;
				var site = context.Site;
				context.Load(site, s => s.Owner);
				context.ExecuteQueryRetry();
				CacheValue = site.Owner.LoginName;
			}
			return CacheValue;
		}
	}
}