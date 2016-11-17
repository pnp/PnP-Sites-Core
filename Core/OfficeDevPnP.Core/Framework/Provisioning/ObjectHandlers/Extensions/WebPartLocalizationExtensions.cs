using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = OfficeDevPnP.Core.Framework.Provisioning.Model.File;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{
    internal static class WebPartLocalizationExtensions
    {
        public static void LocalizeWebParts(this Page page, Web web, TokenParser parser)
        {
            var url = page.Url;
            var webParts = page.WebParts;
            LocalizeParts(web, parser, url, webParts);
        }


        public static void LocalizeWebParts(this File file, Web web, TokenParser parser, Microsoft.SharePoint.Client.File targetFile)
        {
            var url = targetFile.ServerRelativeUrl;
            var webParts = file.WebParts;
            LocalizeParts(web, parser, url, webParts);
        }

        private static void LocalizeParts(Web web, TokenParser parser, string url, WebPartCollection webParts)
        {
            var context = web.Context;
            var allParts = web.GetWebParts(parser.ParseString(url)).ToList();
            foreach (var webPart in webParts)
            {
                var partOnPage = allParts.FirstOrDefault(w => w.ZoneId == webPart.Zone && w.WebPart.ZoneIndex == webPart.Order);
                if (webPart.Title.ContainsResourceToken() && partOnPage != null)
                {
                    var resourceValues = parser.GetResourceTokenResourceValues(webPart.Title);
                    foreach (var resourceValue in resourceValues)
                    {
                        // Save property with correct locale on the request to make it stick
                        // http://sadomovalex.blogspot.no/2015/09/localize-web-part-titles-via-client.html
                        context.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = resourceValue.Item1;
                        partOnPage.WebPart.Properties["Title"] = resourceValue.Item2;
                        partOnPage.SaveWebPartChanges();
                        context.ExecuteQueryRetry();
                    }
                }
            }
            context.PendingRequest.RequestExecutor.WebRequest.Headers.Remove("Accept-Language");
        }
    }
}
