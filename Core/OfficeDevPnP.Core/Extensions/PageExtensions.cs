using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using System.Linq;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using OfficeDevPnP.Core.Utilities.WebParts;
using PersonalizationScope = Microsoft.SharePoint.Client.WebParts.PersonalizationScope;
using System.Net;
using System.IO;
using System.Text;
#if !NETSTANDARD2_0
using System.Web.Configuration;
#endif
using WebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;
using OfficeDevPnP.Core.Pages;
using Microsoft.SharePoint.Client.Utilities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that handles all page and web part related operations
    /// </summary>
    public static partial class PageExtensions
    {
        private const string WikiPage_OneColumn = @"<div class=""ExternalClassC1FD57BEDB8942DC99A06C02F9A98241""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;100%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,1</span></div>";
        private const string WikiPage_OneColumnSideBar = @"<div class=""ExternalClass47565ACDF7974263AA4A556DA974B687""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;66.6%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,2</span></div>";
        private const string WikiPage_TwoColumns = @"<div class=""ExternalClass3811C839E5984CCEA4C8CF738462AFD8""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,2</span></div>";
        private const string WikiPage_TwoColumnsHeader = @"<div class=""ExternalClass850251EB51394304A07A64A05C0BB0F1""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,false,2</span></div>";
        private const string WikiPage_TwoColumnsHeaderFooter = @"<div class=""ExternalClass71C5527252AD45859FA774445D4909A2""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,true,2</span></div>";
        private const string WikiPage_ThreeColumns = @"<div class=""ExternalClass833D1FA704C94892A26C4069C3FE5FE9""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,3</span></div>";
        private const string WikiPage_ThreeColumnsHeader = @"<div class=""ExternalClassD1A150D6187F449B8A6C4BEA2D4913BB""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,false,3</span></div>";
        private const string WikiPage_ThreeColumnsHeaderFooter = @"<div class=""ExternalClass5849C2C61FEC44E9B249C60F7B0ACA38""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,true,3</span></div>";


        /// <summary>
        /// Gets the HTML contents of a wiki page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page, e.g. /sites/demo/SitePages/Test.aspx</param>
        /// <returns>Returns the HTML contents of a wiki page</returns>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl is null</exception>
        public static string GetWikiPageContent(this Web web, string serverRelativePageUrl)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            var file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            web.Context.Load(file, f => f.ListItemAllFields);

            web.Context.ExecuteQueryRetry();

            return file.ListItemAllFields["WikiField"] as string;
        }

        /// <summary>
        /// List the web parts on a page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page containing the webparts</param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl is null</exception>
        public static IEnumerable<WebPartDefinition> GetWebParts(this Web web, string serverRelativePageUrl)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            var file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);
            var limitedWebPartManager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);

            IEnumerable<WebPartDefinition> query;

#if ONPREMISES
            // As long as we've no CSOM library that has the ZoneID we can't use the version check as things don't compile...
            query = web.Context.LoadQuery(limitedWebPartManager.WebParts.IncludeWithDefaultProperties(wp => wp.Id, wp => wp.ZoneId, wp => wp.WebPart, wp => wp.WebPart.Title, wp => wp.WebPart.Properties, wp => wp.WebPart.Hidden));
#else
            if (web.Context.HasMinimalServerLibraryVersion(Constants.MINIMUMZONEIDREQUIREDSERVERVERSION))
            {
                query = web.Context.LoadQuery(limitedWebPartManager.WebParts.IncludeWithDefaultProperties(wp => wp.Id, wp => wp.ZoneId, wp => wp.WebPart, wp => wp.WebPart.Title, wp => wp.WebPart.Properties, wp => wp.WebPart.Hidden));
            }
            else
            {
                query = web.Context.LoadQuery(limitedWebPartManager.WebParts.IncludeWithDefaultProperties(wp => wp.Id, wp => wp.WebPart, wp => wp.WebPart.Title, wp => wp.WebPart.Properties, wp => wp.WebPart.Hidden));
            }
#endif
            web.Context.ExecuteQueryRetry();

            return query;
        }

        /// <summary>
        /// Inserts a web part on a web part page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="webPart">Information about the web part to insert</param>
        /// <param name="page">Page to add the web part on</param>
        /// <returns>Returns the added <see cref="Microsoft.SharePoint.Client.WebParts.WebPartDefinition"/> object</returns>
        /// <exception cref="System.ArgumentException">Thrown when page is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when webPart or page is null</exception>
        public static WebPartDefinition AddWebPartToWebPartPage(this Web web, WebPartEntity webPart, string page)
        {
            if (webPart == null)
            {
                throw new ArgumentNullException(nameof(webPart));
            }

            if (string.IsNullOrEmpty(page))
            {
                throw (page == null)
                  ? new ArgumentNullException(nameof(page))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(page));
            }

            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }
            var serverRelativeUrl = UrlUtility.Combine(web.ServerRelativeUrl, page);

            return AddWebPartToWebPartPage(web, serverRelativeUrl, webPart);
        }

        /// <summary>
        /// Inserts a web part on a web part page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">Page to add the web part on</param>
        /// <param name="webPart">Information about the web part to insert</param>
        /// <returns>Returns the added <see cref="Microsoft.SharePoint.Client.WebParts.WebPartDefinition"/> object</returns>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl or webPart is null</exception>
        public static WebPartDefinition AddWebPartToWebPartPage(this Web web, string serverRelativePageUrl, WebPartEntity webPart)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            if (webPart == null)
            {
                throw new ArgumentNullException(nameof(webPart));
            }

            var webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            if (webPartPage == null)
            {
                return null;
            }

            web.Context.Load(webPartPage);
            web.Context.ExecuteQueryRetry();

            return AddWebPart(web, webPartPage, webPart, webPart.WebPartZone, webPart.WebPartIndex);
        }

        /// <summary>
        /// Add web part to a wiki style page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="folder">System name of the wiki page library - typically sitepages</param>
        /// <param name="webPart">Information about the web part to insert</param>
        /// <param name="page">Page to add the web part on</param>
        /// <param name="row">Row of the wiki table that should hold the inserted web part</param>
        /// <param name="col">Column of the wiki table that should hold the inserted web part</param>
        /// <param name="addSpace">Does a blank line need to be added after the web part (to space web parts)</param>
        /// <returns>Returns the added <see cref="Microsoft.SharePoint.Client.WebParts.WebPartDefinition"/> object</returns>
        /// <exception cref="System.ArgumentException">Thrown when folder or page is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when folder, webPart or page is null</exception>
        public static WebPartDefinition AddWebPartToWikiPage(this Web web, string folder, WebPartEntity webPart, string page, int row, int col, bool addSpace)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw (folder == null)
                  ? new ArgumentNullException(nameof(folder))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(folder));
            }

            if (webPart == null)
            {
                throw new ArgumentNullException(nameof(webPart));
            }

            if (string.IsNullOrEmpty(page))
            {
                throw (page == null)
                  ? new ArgumentNullException(nameof(page))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(page));
            }

            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }

            var webServerRelativeUrl = UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl);
            var serverRelativeUrl = UrlUtility.Combine(folder, page);
            return AddWebPartToWikiPage(web, webServerRelativeUrl + serverRelativeUrl, webPart, row, col, addSpace);
        }

        /// <summary>
        /// Add web part to a wiki style page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to add the webpart to</param>
        /// <param name="webPart">Information about the web part to insert</param>
        /// <param name="row">Row of the wiki table that should hold the inserted web part</param>
        /// <param name="col">Column of the wiki table that should hold the inserted web part</param>
        /// <param name="addSpace">Does a blank line need to be added after the web part (to space web parts)</param>
        /// <returns>Returns the added <see cref="Microsoft.SharePoint.Client.WebParts.WebPartDefinition"/> object</returns>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl or webPart is null</exception>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Xml.XmlDocument.CreateTextNode(System.String)")]
        public static WebPartDefinition AddWebPartToWikiPage(this Web web, string serverRelativePageUrl, WebPartEntity webPart, int row, int col, bool addSpace)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            if (webPart == null)
            {
                throw new ArgumentNullException(nameof(webPart));
            }

            File webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            if (webPartPage == null)
            {
                return null;
            }

            web.Context.Load(webPartPage, wp => wp.ListItemAllFields);
            web.Context.ExecuteQueryRetry();

            string wikiField = (string)webPartPage.ListItemAllFields["WikiField"];

            var wpdNew = AddWebPart(web, webPartPage, webPart, "wpz", 0);

            //HTML structure in default team site home page (W16)
            //<div class="ExternalClass284FC748CB4242F6808DE69314A7C981">
            //  <div class="ExternalClass5B1565E02FCA4F22A89640AC10DB16F3">
            //    <table id="layoutsTable" style="width&#58;100%;">
            //      <tbody>
            //        <tr style="vertical-align&#58;top;">
            //          <td colspan="2">
            //            <div class="ms-rte-layoutszone-outer" style="width&#58;100%;">
            //              <div class="ms-rte-layoutszone-inner" style="word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;">
            //                <div><span><span><div class="ms-rtestate-read ms-rte-wpbox"><div class="ms-rtestate-read 9ed0c0ac-54d0-4460-9f1c-7e98655b0847" id="div_9ed0c0ac-54d0-4460-9f1c-7e98655b0847"></div><div class="ms-rtestate-read" id="vid_9ed0c0ac-54d0-4460-9f1c-7e98655b0847" style="display&#58;none;"></div></div></span></span><p> </p></div>
            //                <div class="ms-rtestate-read ms-rte-wpbox">
            //                  <div class="ms-rtestate-read c7a1f9a9-4e27-4aa3-878b-c8c6c87961c0" id="div_c7a1f9a9-4e27-4aa3-878b-c8c6c87961c0"></div>
            //                  <div class="ms-rtestate-read" id="vid_c7a1f9a9-4e27-4aa3-878b-c8c6c87961c0" style="display&#58;none;"></div>
            //                </div>
            //              </div>
            //            </div>
            //          </td>
            //        </tr>
            //        <tr style="vertical-align&#58;top;">
            //          <td style="width&#58;49.95%;">
            //            <div class="ms-rte-layoutszone-outer" style="width&#58;100%;">
            //              <div class="ms-rte-layoutszone-inner" style="word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;">
            //                <div class="ms-rtestate-read ms-rte-wpbox">
            //                  <div class="ms-rtestate-read b55b18a3-8a3b-453f-a714-7e8d803f4d30" id="div_b55b18a3-8a3b-453f-a714-7e8d803f4d30"></div>
            //                  <div class="ms-rtestate-read" id="vid_b55b18a3-8a3b-453f-a714-7e8d803f4d30" style="display&#58;none;"></div>
            //                </div>
            //              </div>
            //            </div>
            //          </td>
            //          <td class="ms-wiki-columnSpacing" style="width&#58;49.95%;">
            //            <div class="ms-rte-layoutszone-outer" style="width&#58;100%;">
            //              <div class="ms-rte-layoutszone-inner" style="word-wrap&#58;break-word;margin&#58;0px;border&#58;0px;">
            //                <div class="ms-rtestate-read ms-rte-wpbox">
            //                  <div class="ms-rtestate-read 0b2f12a4-3ab5-4a59-b2eb-275bbc617f95" id="div_0b2f12a4-3ab5-4a59-b2eb-275bbc617f95"></div>
            //                  <div class="ms-rtestate-read" id="vid_0b2f12a4-3ab5-4a59-b2eb-275bbc617f95" style="display&#58;none;"></div>
            //                </div>
            //              </div>
            //            </div>
            //          </td>
            //        </tr>
            //      </tbody>
            //    </table>
            //    <span id="layoutsData" style="display&#58;none;">true,false,2</span>
            //  </div>
            //</div>

            // Close all BR tags
            var brRegex = new Regex("<br>", RegexOptions.IgnoreCase);

            wikiField = brRegex.Replace(wikiField, "<br/>");

            var xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div?
            var layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement ??
                               xd.SelectSingleNode("div/table") as XmlElement;

            var layoutsZoneInner = layoutsTable.SelectSingleNode($"tbody/tr[{row}]/td[{col}]/div/div") as XmlElement;
            // - space element
            var space = xd.CreateElement("p");
            var text = xd.CreateTextNode(" ");
            space.AppendChild(text);

            // - wpBoxDiv
            var wpBoxDiv = xd.CreateElement("div");
            layoutsZoneInner.AppendChild(wpBoxDiv);

            if (addSpace)
            {
                layoutsZoneInner.AppendChild(space);
            }

            var attribute = xd.CreateAttribute("class");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read ms-rte-wpbox";
            attribute = xd.CreateAttribute("contentEditable");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "false";
            // - div1
            var div1 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div1);
            div1.IsEmpty = false;
            attribute = xd.CreateAttribute("class");
            div1.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read " + wpdNew.Id.ToString("D");
            attribute = xd.CreateAttribute("id");
            div1.Attributes.Append(attribute);
            attribute.Value = "div_" + wpdNew.Id.ToString("D");
            // - div2
            var div2 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div2);
            div2.IsEmpty = false;
            attribute = xd.CreateAttribute("style");
            div2.Attributes.Append(attribute);
            attribute.Value = "display:none";
            attribute = xd.CreateAttribute("id");
            div2.Attributes.Append(attribute);
            attribute.Value = "vid_" + wpdNew.Id.ToString("D");

            var listItem = webPartPage.ListItemAllFields;
            listItem["WikiField"] = xd.OuterXml;
            listItem.Update();
            web.Context.ExecuteQueryRetry();

            return wpdNew;
        }

        /// <summary>
        /// Gets XML string of a Webpart
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="webPartId">The id of the webpart</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to add the xml to</param>
        /// <returns>Returns XML string of a Webpart</returns>
        public static string GetWebPartXml(this Web web, Guid webPartId, string serverRelativePageUrl)
        {
            string webPartXml = null;

            if (webPartId != Guid.Empty)
            {
#if ONPREMISES
                Guid id = Guid.Empty;

                var wp = web.GetWebParts(serverRelativePageUrl).FirstOrDefault(wps => wps.Id == webPartId);
                if (wp != null)
                {
                    id = wp.Id;
                }
                else
                {
                    return null;
                }
                var uri = new Uri(web.Context.Url);
                var serverRelativeUrl = web.EnsureProperty(w => w.ServerRelativeUrl);
                var webUrl = $"{uri.Scheme}://{uri.Host}:{uri.Port}{serverRelativeUrl}";
                var pageUrl = $"{uri.Scheme}://{uri.Host}:{uri.Port}{serverRelativePageUrl}";
                var request = (HttpWebRequest)WebRequest.Create($"{webUrl}/_vti_bin/exportwp.aspx?pageurl={HttpUtility.UrlKeyValueEncode(pageUrl)}&guidstring={id}");

                var cookieCollection = web.Context.GetCookieCollection();

                if (web.Context.Credentials != null)
                {
                    request.Credentials = web.Context.Credentials;
                }
                else if (cookieCollection != null && cookieCollection.Count > 0)
                {
                    if (request.CookieContainer == null)
                    {
                       request.CookieContainer = new CookieContainer();
                    }
                    request.CookieContainer.Add(cookieCollection);
                }
                else
                {
                    request.UseDefaultCredentials = true;
                }
                
                // apparently without a user agent SharePoint 2013 returns a 302 redirect to an error page without returning the actual web part
                request.UserAgent = "Mozilla/5.0 (Windows NT; Windows NT 6.2; de-DE) pnprocks/5.1.19041.1";

                var response = request.GetResponse();
                using (Stream stream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                    webPartXml = reader.ReadToEnd();
                }

#else

                var webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

                bool forceCheckout = false;
                webPartPage.EnsureProperty(wpg => wpg.ListId);
                if (webPartPage.ListId != Guid.Empty)
                {
                    var list = web.Lists.GetById(webPartPage.ListId);
                    web.Context.Load(list, l => l.ForceCheckout);
                    web.Context.ExecuteQueryRetry();
                    forceCheckout = list.ForceCheckout;
                }

                var limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);

                var query =
                    web.Context.LoadQuery(
                        limitedWebPartManager.WebParts.IncludeWithDefaultProperties(w => w.Id, w => w.ZoneId,
                                w => w.WebPart, w => w.WebPart.Title, w => w.WebPart.Properties, w => w.WebPart.Hidden)
                            .Where(w => w.Id == webPartId));

                web.Context.ExecuteQueryRetry();

                if (query.Any())
                {
                    if (forceCheckout)
                    {
                        webPartPage.CheckOut();
                        web.Context.ExecuteQueryRetry();
                    }
                    var wp = query.First();

                    var exportMode = wp.WebPart.ExportMode;
                    var changed = false;
                    if (exportMode != WebParts.WebPartExportMode.All)
                    {
                        wp.WebPart.ExportMode = WebParts.WebPartExportMode.All;
                        wp.SaveWebPartChanges();
                        web.Context.ExecuteQueryRetry();
                        changed = true;
                    }

                    var result = limitedWebPartManager.ExportWebPart(wp.Id);
                    web.Context.ExecuteQueryRetry();
                    webPartXml = result.Value;

                    if (changed)
                    {
                        wp.WebPart.ExportMode = exportMode;
                        wp.SaveWebPartChanges();
                        web.Context.ExecuteQueryRetry();
                    }
                    if (forceCheckout)
                    {
                        webPartPage.UndoCheckOut();
                        web.Context.ExecuteQueryRetry();
                    }
                }
#endif
            }
            return webPartXml;
        }

        /// <summary>
        /// Applies a layout to a wiki page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="layout">Wiki page layout to be applied</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to add the layout to</param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl is null</exception>
        public static void AddLayoutToWikiPage(this Web web, WikiPageLayout layout, string serverRelativePageUrl)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            var html = "";
            switch (layout)
            {
                case WikiPageLayout.OneColumn:
                    html = WikiPage_OneColumn;
                    break;
                case WikiPageLayout.OneColumnSideBar:
                    html = WikiPage_OneColumnSideBar;
                    break;
                case WikiPageLayout.TwoColumns:
                    html = WikiPage_TwoColumns;
                    break;
                case WikiPageLayout.TwoColumnsHeader:
                    html = WikiPage_TwoColumnsHeader;
                    break;
                case WikiPageLayout.TwoColumnsHeaderFooter:
                    html = WikiPage_TwoColumnsHeaderFooter;
                    break;
                case WikiPageLayout.ThreeColumns:
                    html = WikiPage_ThreeColumns;
                    break;
                case WikiPageLayout.ThreeColumnsHeader:
                    html = WikiPage_ThreeColumnsHeader;
                    break;
                case WikiPageLayout.ThreeColumnsHeaderFooter:
                    html = WikiPage_ThreeColumnsHeaderFooter;
                    break;
                default:
                    break;
            }

            web.AddHtmlToWikiPage(serverRelativePageUrl, html);
        }

        /// <summary>
        /// Applies a layout to a wiki page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="folder">System name of the wiki page library - typically sitepages</param>
        /// <param name="layout">Wiki page layout to be applied</param>
        /// <param name="page">Name of the page that will get a new wiki page layout</param>
        /// <exception cref="System.ArgumentException">Thrown when folder or page is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when folder or page is null</exception>
        public static void AddLayoutToWikiPage(this Web web, string folder, WikiPageLayout layout, string page)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw (folder == null)
                  ? new ArgumentNullException(nameof(folder))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(folder));
            }

            if (string.IsNullOrEmpty(page))
            {
                throw (page == null)
                  ? new ArgumentNullException(nameof(page))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(page));
            }

            var html = "";
            switch (layout)
            {
                case WikiPageLayout.OneColumn:
                    html = WikiPage_OneColumn;
                    break;
                case WikiPageLayout.OneColumnSideBar:
                    html = WikiPage_OneColumnSideBar;
                    break;
                case WikiPageLayout.TwoColumns:
                    html = WikiPage_TwoColumns;
                    break;
                case WikiPageLayout.TwoColumnsHeader:
                    html = WikiPage_TwoColumnsHeader;
                    break;
                case WikiPageLayout.TwoColumnsHeaderFooter:
                    html = WikiPage_TwoColumnsHeaderFooter;
                    break;
                case WikiPageLayout.ThreeColumns:
                    html = WikiPage_ThreeColumns;
                    break;
                case WikiPageLayout.ThreeColumnsHeader:
                    html = WikiPage_ThreeColumnsHeader;
                    break;
                case WikiPageLayout.ThreeColumnsHeaderFooter:
                    html = WikiPage_ThreeColumnsHeaderFooter;
                    break;
                default:
                    break;
            }

            web.AddHtmlToWikiPage(folder, html, page);
        }

        /// <summary>
        /// Add html to a wiki style page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="folder">System name of the wiki page library - typically sitepages</param>
        /// <param name="html">The html to insert</param>
        /// <param name="page">Page to add the html on</param>
        /// <exception cref="System.ArgumentException">Thrown when folder, html or page is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when folder, html or page is null</exception>
        public static void AddHtmlToWikiPage(this Web web, string folder, string html, string page)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw (folder == null)
                  ? new ArgumentNullException(nameof(folder))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(folder));
            }

            if (string.IsNullOrEmpty(html))
            {
                throw (html == null)
                  ? new ArgumentNullException(nameof(html))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(html));
            }

            if (string.IsNullOrEmpty(page))
            {
                throw (page == null)
                  ? new ArgumentNullException(nameof(page))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(page));
            }

            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }

            var webServerRelativeUrl = UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl);

            var serverRelativeUrl = UrlUtility.Combine(webServerRelativeUrl, folder, page);

            AddHtmlToWikiPage(web, serverRelativeUrl, html);
        }

        /// <summary>
        /// Add HTML to a wiki page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to add html to</param>
        /// <param name="html"></param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl or html is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl or html is null</exception>
        public static void AddHtmlToWikiPage(this Web web, string serverRelativePageUrl, string html)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            if (string.IsNullOrEmpty(html))
            {
                throw (html == null)
                  ? new ArgumentNullException(nameof(html))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(html));
            }

            var file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            web.Context.Load(file, f => f.ListItemAllFields);
            web.Context.ExecuteQueryRetry();

            var item = file.ListItemAllFields;

            web.EnsureProperty(w => w.WebTemplate);
            if (web.WebTemplate == "ENTERWIKI")
            {
                item["PublishingPageContent"] = html;
            }
            else
            {
                item["WikiField"] = html;
            }

            item.Update();

            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Add a HTML fragment to a location on a wiki style page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="folder">System name of the wiki page library - typically sitepages</param>
        /// <param name="html">html to be inserted</param>
        /// <param name="page">Page to add the web part on</param>
        /// <param name="row">Row of the wiki table that should hold the inserted web part</param>
        /// <param name="col">Column of the wiki table that should hold the inserted web part</param>
        /// <exception cref="System.ArgumentException">Thrown when folder, html or page is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when folder, html or page is null</exception>
        public static void AddHtmlToWikiPage(this Web web, string folder, string html, string page, int row, int col)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw (folder == null)
                  ? new ArgumentNullException(nameof(folder))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(folder));
            }

            if (string.IsNullOrEmpty(html))
            {
                throw (html == null)
                  ? new ArgumentNullException(nameof(html))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(html));
            }

            if (string.IsNullOrEmpty(page))
            {
                throw (page == null)
                  ? new ArgumentNullException(nameof(page))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(page));
            }

            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }

            var webServerRelativeUrl = UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl);

            var serverRelativeUrl = UrlUtility.Combine(webServerRelativeUrl, folder, page);

            AddHtmlToWikiPage(web, serverRelativeUrl, html, row, col);
        }

        /// <summary>
        /// Add a HTML fragment to a location on a wiki style page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">server relative Url of the page to add the fragment to</param>
        /// <param name="html">html to be inserted</param>
        /// <param name="row">Row of the wiki table that should hold the inserted web part</param>
        /// <param name="col">Column of the wiki table that should hold the inserted web part</param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl or html is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl or html is null</exception>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Xml.XmlDocument.CreateTextNode(System.String)")]
        public static void AddHtmlToWikiPage(this Web web, string serverRelativePageUrl, string html, int row, int col)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            if (string.IsNullOrEmpty(html))
            {
                throw (html == null)
                  ? new ArgumentNullException(nameof(html))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(html));
            }

            var file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            web.Context.Load(file, f => f.ListItemAllFields);
            web.Context.ExecuteQueryRetry();

            var item = file.ListItemAllFields;

            var wikiField = (string)item["WikiField"];

            var xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div?
            var layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement;
            if (layoutsTable == null)
            {
                layoutsTable = xd.SelectSingleNode("div/table") as XmlElement;
            }

            // Add the html content
            var layoutsZoneInner = layoutsTable.SelectSingleNode($"tbody/tr[{row}]/td[{col}]/div/div") as XmlElement;
            var text = xd.CreateTextNode("!!123456789!!");
            layoutsZoneInner.AppendChild(text);

            item["WikiField"] = xd.OuterXml.Replace("!!123456789!!", html);
            item.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Deletes a web part from a page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="folder">System name of the wiki page library - typically sitepages</param>
        /// <param name="title">Title of the web part that needs to be deleted</param>
        /// <param name="page">Page to remove the web part from</param>
        /// <exception cref="System.ArgumentException">Thrown when folder, title or page is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when folder, title or page is null</exception>
        public static void DeleteWebPart(this Web web, string folder, string title, string page)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw (folder == null)
                  ? new ArgumentNullException(nameof(folder))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(folder));
            }

            if (string.IsNullOrEmpty(title))
            {
                throw (title == null)
                  ? new ArgumentNullException(nameof(title))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(title));
            }

            if (string.IsNullOrEmpty(page))
            {
                throw (page == null)
                  ? new ArgumentNullException(nameof(page))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(page));
            }

            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }

            var webServerRelativeUrl = UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl);

            var serverRelativeUrl = UrlUtility.Combine(folder, page);

            DeleteWebPart(web, webServerRelativeUrl + serverRelativeUrl, title);
        }

        /// <summary>
        /// Deletes a web part from a page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to remove</param>
        /// <param name="title">Title of the web part that needs to be deleted</param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl or title is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl or title is null</exception>
        public static void DeleteWebPart(this Web web, string serverRelativePageUrl, string title)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            if (string.IsNullOrEmpty(title))
            {
                throw (title == null)
                  ? new ArgumentNullException(nameof(title))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(title));
            }

            var webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            if (webPartPage == null)
            {
                return;
            }

            web.Context.Load(webPartPage);
            web.Context.ExecuteQueryRetry();

            var limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            web.Context.Load(limitedWebPartManager.WebParts, wps => wps.Include(wp => wp.WebPart.Title));
            web.Context.ExecuteQueryRetry();

            if (limitedWebPartManager.WebParts.Count >= 0)
            {
                for (var i = 0; i < limitedWebPartManager.WebParts.Count; i++)
                {
                    var oWebPart = limitedWebPartManager.WebParts[i].WebPart;
                    if (oWebPart.Title.Equals(title, StringComparison.InvariantCultureIgnoreCase))
                    {
                        limitedWebPartManager.WebParts[i].DeleteWebPart();
                        web.Context.ExecuteQueryRetry();
                        break;
                    }
                }
            }
        }
        
#if !SP2013 && !SP2016
        /// <summary>
        /// Adds a client side "modern" page to a "classic" or "modern" site
        /// </summary>
        /// <param name="web">Web to add the page to</param>
        /// <param name="pageName">Name (e.g. demo.aspx) of the page to be added</param>
        /// <param name="alreadyPersist">Already persist the created, empty, page before returning the instantiated <see cref="ClientSidePage"/> instance</param>
        /// <returns>A <see cref="ClientSidePage"/> instance</returns>
        public static ClientSidePage AddClientSidePage(this Web web, string pageName = "", bool alreadyPersist = false)
        {
            var page = new ClientSidePage(web.Context as ClientContext);

            if (alreadyPersist)
            {
                page.Save(pageName);
            }
            return page;
        }
        
        /// <summary>
        /// Loads a client side "modern" page
        /// </summary>
        /// <param name="web">Web to load the page from</param>
        /// <param name="pageName">Name (e.g. demo.aspx) of the page to be loaded</param>
        /// <returns>A <see cref="ClientSidePage"/> instance</returns>
        public static ClientSidePage LoadClientSidePage(this Web web, string pageName)
        {
            return ClientSidePage.Load((web.Context as ClientContext), pageName);
        }
#endif

        /// <summary>
        /// Adds a blank Wiki page to the site pages library
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="wikiPageLibraryName">Name of the wiki page library</param>
        /// <param name="wikiPageName">Wiki page to operate on</param>
        /// <returns>The relative URL of the added wiki page</returns>
        /// <exception cref="System.ArgumentException">Thrown when wikiPageLibraryName or wikiPageName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when wikiPageLibraryName or wikiPageName is null</exception>
        public static string AddWikiPage(this Web web, string wikiPageLibraryName, string wikiPageName)
        {
            string wikiPageUrl, newWikiPageUrl, pathAndQuery;
            List pageLibrary;
            File currentPageFile;

            WikiPageImplementation(web, wikiPageLibraryName, wikiPageName, out wikiPageUrl, out pageLibrary, out newWikiPageUrl, out currentPageFile, out pathAndQuery);

            if (!currentPageFile.Exists)
            {
                var newpage = pageLibrary.RootFolder.Files.AddTemplateFile(newWikiPageUrl, TemplateFileType.WikiPage);
                web.Context.Load(newpage);
                web.Context.ExecuteQueryRetry();
                wikiPageUrl = newpage.ServerRelativeUrl.Replace(pathAndQuery, "");
            }

            return wikiPageUrl;
        }

        /// <summary>
        /// Returns the Url for the requested wiki page, creates it if the pageis not yet available
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="wikiPageLibraryName">Name of the wiki page library</param>
        /// <param name="wikiPageName">Wiki page to operate on</param>
        /// <returns>The relative URL of the added wiki page</returns>
        /// <exception cref="System.ArgumentException">Thrown when wikiPageLibraryName or wikiPageName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when wikiPageLibraryName or wikiPageName is null</exception>
        public static string EnsureWikiPage(this Web web, string wikiPageLibraryName, string wikiPageName)
        {
            string wikiPageUrl, newWikiPageUrl, pathAndQuery;
            List pageLibrary;
            File currentPageFile;

            WikiPageImplementation(web, wikiPageLibraryName, wikiPageName, out wikiPageUrl, out pageLibrary, out newWikiPageUrl, out currentPageFile, out pathAndQuery);

            if (!currentPageFile.Exists)
            {
                var newpage = pageLibrary.RootFolder.Files.AddTemplateFile(newWikiPageUrl, TemplateFileType.WikiPage);
                web.Context.Load(newpage);
                web.Context.ExecuteQueryRetry();
                wikiPageUrl = newpage.ServerRelativeUrl.Replace(pathAndQuery, "");
            }
            else
            {
                web.Context.Load(currentPageFile, s => s.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
                wikiPageUrl = currentPageFile.ServerRelativeUrl.Replace(pathAndQuery, "");
            }

            return wikiPageUrl;
        }

        private static void WikiPageImplementation(Web web, string wikiPageLibraryName, string wikiPageName, out string wikiPageUrl, out List pageLibrary, out string newWikiPageUrl, out File currentPageFile, out string pathAndQuery)
        {
            if (string.IsNullOrEmpty(wikiPageLibraryName))
            {
                throw (wikiPageLibraryName == null)
                  ? new ArgumentNullException(nameof(wikiPageLibraryName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(wikiPageLibraryName));
            }

            if (string.IsNullOrEmpty(wikiPageName))
            {
                throw (wikiPageName == null)
                  ? new ArgumentNullException(nameof(wikiPageName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(wikiPageName));
            }

            wikiPageUrl = "";
            pageLibrary = web.Lists.GetByTitle(wikiPageLibraryName);
            web.Context.Load(web, w => w.Url);
            web.Context.Load(pageLibrary.RootFolder, f => f.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            var pageLibraryUrl = pageLibrary.RootFolder.ServerRelativeUrl;
            newWikiPageUrl = pageLibraryUrl + "/" + wikiPageName;
            currentPageFile = web.GetFileByServerRelativeUrl(newWikiPageUrl);
            web.Context.Load(currentPageFile, f => f.Exists);
            web.Context.ExecuteQueryRetry();

            pathAndQuery = new Uri(web.Url).PathAndQuery;
            if (!pathAndQuery.EndsWith("/"))
            {
                pathAndQuery = pathAndQuery + "/";
            }
        }

        /// <summary>
        /// Adds a wiki page by Url
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the wiki page to process</param>
        /// <param name="html">HTML to add to wiki page</param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl is null</exception>
        public static void AddWikiPageByUrl(this Web web, string serverRelativePageUrl, string html = null)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            web.EnsureProperties(w => w.ServerRelativeUrl, w => w.WebTemplate);

            var folderName = serverRelativePageUrl.Substring(0, serverRelativePageUrl.LastIndexOf("/", StringComparison.Ordinal));

            //ensure that folderName does not contain the web's ServerRelativeUrl -> otherwise it will fail on SubSites
            if (folderName.ToLower().StartsWith((web.ServerRelativeUrl.ToLower())))
            {
                folderName = folderName.Substring(web.ServerRelativeUrl.Length);
            }
            var folder = web.EnsureFolderPath(folderName);

            if (web.WebTemplate == "ENTERWIKI")
            {
                if(!serverRelativePageUrl.StartsWith("/"))
                {
                    serverRelativePageUrl = UrlUtility.Combine(web.ServerRelativeUrl, serverRelativePageUrl);
                }
                var filename = serverRelativePageUrl.Substring(serverRelativePageUrl.LastIndexOf("/")+1);
                web.AddPublishingPage(filename, "EnterpriseWiki", null, folder: folder);
                var file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);
                file.ListItemAllFields["PublishingPageContent"] = html;
                file.ListItemAllFields.Update();
                file.ListItemAllFields.Context.ExecuteQueryRetry();
            }
            else
            {
                folder.Files.AddTemplateFile(serverRelativePageUrl, TemplateFileType.WikiPage);
                web.Context.ExecuteQueryRetry();
                if (html != null)
                {
                    web.AddHtmlToWikiPage(serverRelativePageUrl, html);
                }
            }          
        }

        /// <summary>
        /// Sets a web part property
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="key">The key to update</param>
        /// <param name="value">The value to set</param>
        /// <param name="id">The id of the webpart</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to set web part property</param>
        /// <exception cref="System.ArgumentException">Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when key or serverRelativePageUrl is null</exception>
        public static void SetWebPartProperty(this Web web, string key, string value, Guid id, string serverRelativePageUrl)
        {
            SetWebPartPropertyInternal(web, key, value, id, serverRelativePageUrl);
        }

        /// <summary>
        /// Sets a web part property
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="key">The key to update</param>
        /// <param name="value">The value to set</param>
        /// <param name="id">The id of the webpart</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to set web part property</param>
        /// <exception cref="System.ArgumentException">Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when key or serverRelativePageUrl is null</exception>
        public static void SetWebPartProperty(this Web web, string key, int value, Guid id, string serverRelativePageUrl)
        {
            SetWebPartPropertyInternal(web, key, value, id, serverRelativePageUrl);
        }

        /// <summary>
        /// Sets a web part property
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="key">The key to update</param>
        /// <param name="value">The value to set</param>
        /// <param name="id">The id of the webpart</param>
        /// <param name="serverRelativePageUrl">Server relative URL of the page to set web part property</param>
        /// <exception cref="System.ArgumentException">Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when key or serverRelativePageUrl is null</exception>
        public static void SetWebPartProperty(this Web web, string key, bool value, Guid id, string serverRelativePageUrl)
        {
            SetWebPartPropertyInternal(web, key, value, id, serverRelativePageUrl);
        }

        private static WebPartDefinition AddWebPart(Web fileWeb, File webPartPage, WebPartEntity webPart, string zoneId, int zoneIndex)
        {
            var limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var oWebPartDefinition = limitedWebPartManager.ImportWebPart(webPart.WebPartXml);

            var wpdNew = limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, zoneId, zoneIndex);
            webPartPage.Context.Load(wpdNew);
            webPartPage.Context.ExecuteQueryRetry();

            var webPartPostProcessor = WebPartPostProcessorFactory.Resolve(webPart.WebPartXml);

            var currentContext = ((ClientContext)webPartPage.Context);
            var contextWeb = currentContext.Web;

            contextWeb.EnsureProperties(w => w.Url, w => w.Id);
            fileWeb.EnsureProperties(w => w.Url, w => w.Id);

            if (contextWeb.Id.Equals(fileWeb.Id))
            {
                webPartPostProcessor.Process(wpdNew, webPartPage);
            }
            else
            {
                using (var context = currentContext.Clone(fileWeb.Url))
                {
#if !SP2013
                    webPartPage.EnsureProperties(f => f.UniqueId);
                    var file = context.Web.GetFileById(webPartPage.UniqueId);
#else
                    webPartPage.EnsureProperties(f => f.ServerRelativeUrl);
                    var file = context.Web.GetFileByServerRelativeUrl(webPartPage.ServerRelativeUrl);
#endif
                    webPartPostProcessor.Process(wpdNew, file);
                }
            }

            return wpdNew;
        }

        private static void SetWebPartPropertyInternal(this Web web, string key, object value, Guid id, string serverRelativePageUrl)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw (key == null)
                  ? new ArgumentNullException(nameof(key))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(key));
            }

            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            var context = web.Context as ClientContext;

            var file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            context.Load(file);
            context.ExecuteQueryRetry();

            var wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);

            context.Load(wpm.WebParts);

            context.ExecuteQueryRetry();

            var def = wpm.WebParts.GetById(id);

            context.Load(def);
            context.ExecuteQueryRetry();

            switch (key.ToLower())
            {
                case "title":
                    {
                        def.WebPart.Title = value as string;
                        break;
                    }
                case "titleurl":
                    {
                        def.WebPart.TitleUrl = value as string;
                        break;
                    }
                default:
                    {
                        def.WebPart.Properties[key] = value;
                        break;
                    }
            }


            def.SaveWebPartChanges();

            context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Returns web part properties
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="id">The id of the webpart</param>
        /// <param name="serverRelativePageUrl"></param>
        /// <exception cref="System.ArgumentException">Thrown when key or serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when key or serverRelativePageUrl is null</exception>
        public static PropertyValues GetWebPartProperties(this Web web, Guid id, string serverRelativePageUrl)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException(nameof(serverRelativePageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(serverRelativePageUrl));
            }

            var context = web.Context as ClientContext;

            var file = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            context.Load(file);
            context.ExecuteQueryRetry();

            var wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);

            var def = wpm.WebParts.GetById(id);

            context.Load(def.WebPart.Properties);
            context.ExecuteQueryRetry();

            return def.WebPart.Properties;
        }

        /// <summary>
        /// Adds a user-friendly URL for a PublishingPage object.
        /// </summary>
        /// <param name="page">The target page to add to managed navigation.</param>
        /// <param name="web">The target web.</param>
        /// <param name="navigationTitle">The title for the navigation item.</param>
        /// <param name="friendlyUrlSegment">The user-friendly text to use as the URL segment.</param>
        /// <param name="editableParent">The parent NavigationTermSetItem object below which this new friendly URL should be created.</param>
        /// <param name="showInGlobalNavigation">Defines whether the navigation item has to be shown in the Global Navigation, optional and default to true.</param>
        /// <param name="showInCurrentNavigation">Defines whether the navigation item has to be shown in the Current Navigation, optional and default to true.</param>
        /// <returns>The simple link URL just created.</returns>
        public static string AddNavigationFriendlyUrl(this PublishingPage page, Web web,
            string navigationTitle, string friendlyUrlSegment, NavigationTermSetItem editableParent,
            bool showInGlobalNavigation = true, bool showInCurrentNavigation = true)
        {
            // Add the Friendly URL
            var friendlyUrl = page.AddFriendlyUrl(friendlyUrlSegment, editableParent, true);

            // Retrieve terms for searching parent
            web.Context.Load(editableParent.Terms, ts => ts.Include(t => t.FriendlyUrlSegment, t => t.Title));
            web.Context.ExecuteQueryRetry();

            // Configure the friendly URL Title
            var friendlyUrlTerm = editableParent.Terms
                .FirstOrDefault(t => t.FriendlyUrlSegment.Value == friendlyUrlSegment);
            if (friendlyUrlTerm != null)
            {
                // Assign term label for taxonomy equal to navigation item title
                var friendlyUrlBackingTerm = friendlyUrlTerm.GetTaxonomyTerm();
                web.Context.Load(friendlyUrlBackingTerm, t => t.Labels);
                web.Context.ExecuteQueryRetry();

                var defaultLabel = friendlyUrlBackingTerm.Labels.FirstOrDefault(l => l.IsDefaultForLanguage);
                if (defaultLabel != null)
                {
                    defaultLabel.Value = navigationTitle;
                }

                // Configure the navigation settings
                if (!showInGlobalNavigation || !showInCurrentNavigation)
                {
                    if (!showInGlobalNavigation && !showInCurrentNavigation)
                    {
                        friendlyUrlBackingTerm.SetLocalCustomProperty("_Sys_Nav_ExcludedProviders", "\"GlobalNavigationTaxonomyProvider\",\"CurrentNavigationTaxonomyProvider\"");
                    }
                    if (!showInGlobalNavigation)
                    {
                        friendlyUrlBackingTerm.SetLocalCustomProperty("_Sys_Nav_ExcludedProviders", "\"GlobalNavigationTaxonomyProvider\"");
                    }
                    else
                    {
                        friendlyUrlBackingTerm.SetLocalCustomProperty("_Sys_Nav_ExcludedProviders", "\"CurrentNavigationTaxonomyProvider\"");
                    }
                }

                // Assign term title for site navigation
                friendlyUrlTerm.Title.Value = navigationTitle;
                friendlyUrlTerm.GetTaxonomyTermStore().CommitAll();
                web.Context.ExecuteQueryRetry();
            }

            return (friendlyUrl.Value);
        }


    }
}
