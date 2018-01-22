using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client.Publishing;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for publishing extension methods
    /// </summary>
    public static class PublishingExtensions
    {
        #region Publishing Pages
        /// <summary>
        /// Adds the publishing page.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="pageName">Name of the page.</param>
        /// <param name="pageTemplateName">Name of the page template/layout excluded the .aspx file extension.</param>
        /// <param name="title">The title of the target publishing page.</param>
        /// <param name="publish">Should the page be published or not?</param>
        /// <param name="folder">The target folder for the page, within the Pages library.</param>
        /// <param name="startDate">Start date for scheduled publishing.</param>
        /// <param name="endDate">End date for scheduled publishing.</param>
        /// <param name="schedule">Defines whether to define a schedule or not.</param>
        /// <exception cref="System.ArgumentNullException">Thrown when key or pageName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentException">Thrown when key or pageName is null</exception>
        public static void AddPublishingPage(this Web web, string pageName, string pageTemplateName, string title = null, bool publish = false, Folder folder = null, DateTime? startDate = null, DateTime? endDate = null, Boolean schedule = false)
        {
            if (string.IsNullOrEmpty(pageName))
            {
                throw (title == null)
                  ? new ArgumentNullException(nameof(pageName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(pageName));
            }
            if (string.IsNullOrEmpty(pageTemplateName))
            {
                throw (title == null)
                  ? new ArgumentNullException(nameof(pageTemplateName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(pageTemplateName));
            }
            if (string.IsNullOrEmpty(title))
            {
                title = pageName;
            }

            // Fix page name, if needed
            pageName = pageName.ReplaceInvalidUrlChars("-");

            var context = web.Context as ClientContext;
            var site = context.Site;
            context.Load(site, s => s.ServerRelativeUrl);
            context.ExecuteQueryRetry();

            // Load reference Page Layout
            var pageFromPageLayout = context.Site.RootWeb.GetFileByServerRelativeUrl($"{UrlUtility.EnsureTrailingSlash(site.ServerRelativeUrl)}_catalogs/masterpage/{pageTemplateName}.aspx");
            var pageLayoutItem = pageFromPageLayout.ListItemAllFields;
            context.Load(pageLayoutItem);
            context.ExecuteQueryRetry();

            // Create the publishing page
            var publishingWeb = PublishingWeb.GetPublishingWeb(context, web);
            context.Load(publishingWeb);

            // Configure the publishing page
            var pageInformation = new PublishingPageInformation
            {
                Name = !pageName.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase) ?
                    $"{pageName}.aspx" : pageName,
                PageLayoutListItem = pageLayoutItem
            };

            // Handle target folder, if any
            if (folder != null)
            {
                pageInformation.Folder = folder;
            }
            var page = publishingWeb.AddPublishingPage(pageInformation);

            // Get parent list of item, this way we can handle all languages
            var pagesLibrary = page.ListItem.ParentList;
            context.Load(pagesLibrary);
            context.ExecuteQueryRetry();
            var pageItem = page.ListItem;
            pageItem["Title"] = title;
            pageItem.Update();

            // Checkin the page file, if needed
            web.Context.Load(pageItem, p => p.File.CheckOutType);
            web.Context.ExecuteQueryRetry();
            if (pageItem.File.CheckOutType != CheckOutType.None)
            {
                pageItem.File.CheckIn(string.Empty, CheckinType.MajorCheckIn);
            }

            // Publish the page, if required
            if (publish)
            {
                pageItem.File.Publish(string.Empty);
                if (pagesLibrary.EnableModeration)
                {
                    pageItem.File.Approve(string.Empty);

                    // Setup scheduling, if required
                    if (schedule && startDate.HasValue)
                    {
                        page.StartDate = startDate.Value;
                        page.EndDate = endDate ?? new DateTime(2050, 01, 01);
                        page.Schedule(string.Empty);
                    }
                }
            }
            context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Gets a Publishing Page from the root folder of the Pages library.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="fileLeafRef">The file leaf reference.</param>
        /// <returns>The PublishingPage object, if any. Otherwise null.</returns>
        /// <exception cref="System.ArgumentNullException">fileLeafRef</exception>
        /// <exception cref="System.ArgumentException">fileLeafRef</exception>
        public static PublishingPage GetPublishingPage(this Web web, string fileLeafRef)
        {
            return (web.GetPublishingPage(fileLeafRef, null));
        }

        /// <summary>
        /// Gets a Publishing Page from any folder in the Pages library.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="fileLeafRef">The file leaf reference.</param>
        /// <param name="folder">The folder where to search the page.</param>
        /// <returns>The PublishingPage object, if any. Otherwise null.</returns>
        /// <exception cref="System.ArgumentNullException">fileLeafRef</exception>
        /// <exception cref="System.ArgumentException">fileLeafRef</exception>
        public static PublishingPage GetPublishingPage(this Web web, string fileLeafRef, Folder folder)
        {
            if (string.IsNullOrEmpty(fileLeafRef))
            {
                throw (fileLeafRef == null)
                  ? new ArgumentNullException(nameof(fileLeafRef))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(fileLeafRef));
            }

            var context = web.Context as ClientContext;
            var pages = web.GetPagesLibrary();
            // Get the language agnostic "Pages" library name         
            context.Load(pages, p => p.RootFolder, p => p.ItemCount);
            context.ExecuteQueryRetry();

            if (pages != null && pages.ItemCount > 0)
            {
                var camlQuery = new CamlQuery
                {
                    FolderServerRelativeUrl = folder != null ? folder.ServerRelativeUrl : pages.RootFolder.ServerRelativeUrl,
                    ViewXml = $@"<View Scope='RecursiveAll'>  
                                    <Query> 
                                        <Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='Text'>{fileLeafRef}</Value></Eq></Where> 
                                    </Query> 
                                </View>"
                };

                var listItems = pages.GetItems(camlQuery);
                context.Load(listItems);
                context.ExecuteQueryRetry();

                if (listItems.Count > 0)
                {
                    var page = PublishingPage.GetPublishingPage(context, listItems[0]);
                    context.Load(page);
                    context.ExecuteQueryRetry();
                    return page;
                }
            }

            return null;
        }
        #endregion

        #region Image Renditions

        /// <summary>
        /// Creates an Image Rendition if the name of the Image Rendition doesn't already exist.
        /// </summary>
        /// <param name="web">SharePoint Web</param>
        /// <param name="imageRenditionName">The display name of the Image Rendition</param>
        /// <param name="imageRenditionWidth">The width of the Image Rendition</param>
        /// <param name="imageRenditionHeight">The height of the Image Rendition</param>
        public static void CreatePublishingImageRendition(this Web web, string imageRenditionName, int imageRenditionWidth, int imageRenditionHeight)
        {
            List<string> imageRenditionNames = new List<string>();
            List<ImageRendition> existingImageRenditions = SiteImageRenditions.GetRenditions(web.Context) as List<ImageRendition>;
            web.Context.ExecuteQueryRetry();
            foreach (ImageRendition existingImageRendition in existingImageRenditions)
            {
                imageRenditionNames.Add(existingImageRendition.Name);
            }
            if (!imageRenditionNames.Contains(imageRenditionName))
            {
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_CreatePublishingImageRendition, imageRenditionName, imageRenditionWidth, imageRenditionHeight);
                ImageRendition newImageRendition = new ImageRendition();
                newImageRendition.Name = imageRenditionName;
                newImageRendition.Width = imageRenditionWidth;
                newImageRendition.Height = imageRenditionHeight;
                existingImageRenditions.Add(newImageRendition);
                SiteImageRenditions.SetRenditions(web.Context, existingImageRenditions);
                web.Context.ExecuteQueryRetry();
            }
            else
            {
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.WebExtensions_CreatePublishingImageRendition_Error, imageRenditionName);
            }
        }

        /// <summary>
        /// Removes an existing image rendition
        /// </summary>
        /// <param name="web">SharePoint Web</param>
        /// <param name="imageRenditionName">The name of the image rendition</param>
        public static void RemovePublishingImageRendition(this Web web, string imageRenditionName)
        {
            var imageRenditions = SiteImageRenditions.GetRenditions(web.Context);
            web.Context.ExecuteQueryRetry();
            var newRenditionList = imageRenditions.Where(i => i.Name != imageRenditionName).ToList();
            SiteImageRenditions.SetRenditions(web.Context, newRenditionList);
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Returns all image renditions
        /// </summary>
        /// <param name="web">SharePoint Web</param>
        public static IList<ImageRendition> GetPublishingImageRenditions(this Web web)
        {
            var imageRenditions = SiteImageRenditions.GetRenditions(web.Context);
            web.Context.ExecuteQueryRetry();
            return imageRenditions;
        }
        #endregion
    }
}
