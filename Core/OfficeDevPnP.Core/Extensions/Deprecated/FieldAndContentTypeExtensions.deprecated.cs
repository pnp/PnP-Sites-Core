using System;
using System.ComponentModel;
using System.Linq;
using System.Xml;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// This class provides extension methods that will help you work with fields and content types.
    /// </summary>
    public static partial class FieldAndContentTypeExtensions
    {
        /// <summary>
        /// Set's default content type list. 
        /// </summary>
        /// <remarks>Notice. Currently removes other content types from the list. Known issue</remarks>
        /// <param name="list">List to update</param>
        /// <param name="contentTypeId">Complete ID for the content type</param>
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the August 2018 release.")]
        public static void SetDefaultContentTypeToList(this List list, string contentTypeId)
        {
            var ctCol = list.ContentTypes;
            list.Context.Load(ctCol);
            list.Context.ExecuteQueryRetry();

            var ctIds = ctCol.AsEnumerable().Select(ct => ct.Id).ToList();

            // remove the folder content type
            var newOrder = ctIds.Except(ctIds.Where(id => id.StringValue.StartsWith("0x012000")))
                                 .OrderBy(x => !x.StringValue.StartsWith(contentTypeId, StringComparison.OrdinalIgnoreCase))
                                 .ToArray();
            list.RootFolder.UniqueContentTypeOrder = newOrder;

            list.RootFolder.Update();
            list.Update();
            list.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Set default content type to list
        /// </summary>
        /// <param name="list">List to update</param>
        /// <param name="contentType">Content type to make default</param>
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the August 2018 release.")]
        public static void SetDefaultContentTypeToList(this List list, ContentType contentType)
        {
            SetDefaultContentTypeToList(list, contentType.Id.ToString());
        }

        /// <summary>
        /// Set default content type to list
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="list">List to update</param>
        /// <param name="contentType">Content type to make default</param>
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the August 2018 release.")]
        public static void SetDefaultContentTypeToList(this Web web, List list, ContentType contentType)
        {
            SetDefaultContentTypeToList(web, list, contentType.Id.ToString());
        }

        /// <summary>
        /// Set default content type to list
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="list">List to update</param>
        /// <param name="contentTypeId">Complete ID for the content type</param>        
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the August 2018 release.")]
        public static void SetDefaultContentTypeToList(this Web web, List list, string contentTypeId)
        {
            list.SetDefaultContentTypeToList(contentTypeId);
        }

        /// <summary>
        /// Set default content type to list
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listTitle">Title of the list to be updated</param>
        /// <param name="contentTypeId">Complete ID for the content type</param>
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the August 2018 release.")]
        public static void SetDefaultContentTypeToList(this Web web, string listTitle, string contentTypeId)
        {
            var list = web.GetListByTitle(listTitle);
            web.Context.Load(list);
            web.Context.ExecuteQueryRetry();
            web.SetDefaultContentTypeToList(list, contentTypeId);
        }

        /// <summary>
        /// Set's default content type list. 
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listTitle">Title of the list to be updated</param>
        /// <param name="contentType">Content type to make default</param>
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the August 2018 release.")]
        public static void SetDefaultContentTypeToList(this Web web, string listTitle, ContentType contentType)
        {
            SetDefaultContentTypeToList(web, listTitle, contentType.Id.ToString());
        }

    }
}
