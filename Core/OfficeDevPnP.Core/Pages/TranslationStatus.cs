using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Pages
{
    /// <summary>
    /// Translation status of a page
    /// </summary>
    public sealed class TranslationStatusCollection
    {
        [JsonProperty]
        public List<string> UntranslatedLanguages { get; set; }

        [JsonProperty]
        public List<TranslationStatus> Items { get; set; }
    }

    /// <summary>
    /// Translation page url
    /// </summary>
    public sealed class TranslationPath
    {
        [JsonProperty]
        public string DecodedUrl { get; set; }
    }

    /// <summary>
    /// Translation status for a page
    /// </summary>
    public sealed class TranslationStatus
    {
        /// <summary>
        /// The culture of this translation
        /// </summary>
        [JsonProperty]
        internal string Culture { get;set; }

        /// <summary>
        /// The web-relative path to this translation
        /// </summary>
        [JsonProperty]
        internal TranslationPath Path { get; set; }

        /// <summary>
        /// Last modified date of this translation
        /// </summary>
        [JsonProperty]
        internal DateTime LastModified { get; set; }

        /// <summary>
        /// The file status (checked out, draft, published) of this translation
        /// </summary>
        [JsonProperty]
        internal FileLevel FileStatus { get; set; }

        /// <summary>
        /// The file status (checked out, draft, published) of this translation
        /// </summary>
        [JsonProperty]
        internal bool HasPublishedVersion { get; set; }

        /// <summary>
        /// The page title of this translation
        /// </summary>
        [JsonProperty]
        internal string Title { get; set; }
    }
}
