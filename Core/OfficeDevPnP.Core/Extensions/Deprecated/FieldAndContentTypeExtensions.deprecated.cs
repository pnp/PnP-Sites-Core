using System;
using System.ComponentModel;
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
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the May 2018 release.")]
        public static void SetDefaultContentTypeToList(this List list, string contentTypeId)
        {
            list.SetDefaultContentType(contentTypeId);
        }

        /// <summary>
        /// Set default content type to list
        /// </summary>
        /// <param name="list">List to update</param>
        /// <param name="contentType">Content type to make default</param>
        [Obsolete("Use List.SetDefaultContentType method. This method will be removed in the May 2018 release.")]
        public static void SetDefaultContentTypeToList(this List list, ContentType contentType)
        {
            SetDefaultContentTypeToList(list, contentType.Id.ToString());
        }
    }
}
