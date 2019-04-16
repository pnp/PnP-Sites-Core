using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{sitename}",
      Description = "Returns the title of the current site",
      Example = "{sitename}",
      Returns = "My Company Portal")]
    internal class SiteTitleToken : VolatileTokenDefinition
    {
        public SiteTitleToken(Web web) : base(web, "{sitetitle}", "{sitename}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Load(TokenContext.Web, w => w.Title);
                TokenContext.ExecuteQueryRetry();
                CacheValue = TokenContext.Web.Title;
            }
            return CacheValue;
        }

        /// <summary>
        /// Replaces the specified value with the Site Title token.
        /// </summary>
        /// <param name="value">String value to replace with site title token.</param>
        /// <param name="web">The web used to match <paramref name="value"/> with the web's title.</param>
        /// <returns></returns>
        public static string GetReplaceToken(string value, Web web)
        {
            web.EnsureProperty(w => w.Title);

            if (string.IsNullOrEmpty(web.Title))
            {
                return value;
            }
            else
            {
                return value.Replace(web.Title, "{sitetitle}");
            }
        }
    }
}