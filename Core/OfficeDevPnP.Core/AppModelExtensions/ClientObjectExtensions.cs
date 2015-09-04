using System;
using System.Linq.Expressions;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.AppModelExtensions
{
	public static class ClientObjectExtensions
	{
		public static bool IsPropertyAvailable<T>(this T clientObject, Expression<Func<T, object>> propertySelector) where T : ClientObject
		{
			var body = propertySelector.Body as MemberExpression;

			if (body == null)
			{
				body = ((UnaryExpression)propertySelector.Body).Operand as MemberExpression;
			}

			return clientObject.IsPropertyAvailable(body.Member.Name);
		}

		public static bool IsObjectPropertyInstantiated<T>(this T clientObject, Expression<Func<T, object>> propertySelector) where T : ClientObject
		{
			var body = propertySelector.Body as MemberExpression;

			if (body == null)
			{
				body = ((UnaryExpression)propertySelector.Body).Operand as MemberExpression;
			}

			return clientObject.IsObjectPropertyInstantiated(body.Member.Name);
		}
	}
}
