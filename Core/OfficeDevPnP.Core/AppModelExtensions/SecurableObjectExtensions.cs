using System;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.AppModelExtensions
{
	public static class SecurableObjectExtensions
	{
		public static Web GetAssociatedWeb(this SecurableObject securable)
		{
			if (securable is Web)
			{
				return (Web)securable;
			}

			if (securable is List)
			{
				var list = (List)securable;
				var web = list.ParentWeb;
				securable.Context.Load(web);
				securable.Context.ExecuteQueryRetry();

				return web;
			}

			if (securable is ListItem)
			{
				var listItem = (ListItem)securable;
				var web = listItem.ParentList.ParentWeb;
				securable.Context.Load(web);
				securable.Context.ExecuteQueryRetry();

				return web;
			}

			throw new Exception("Only Web, List, ListItem supported as SecurableObjects");
		}
	}
}
