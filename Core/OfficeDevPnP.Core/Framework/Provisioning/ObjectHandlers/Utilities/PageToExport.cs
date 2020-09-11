using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Class that defines a page to be exported
    /// </summary>
    public class PageToExport
    {
        public string PageName { get; set; }

        public string PageUrl { get; set; }

        public Guid PageId { get; set; }

        public Guid SourcePageId { get; set; }

        public string SourcePageName { get; set; }

        public ListItem ListItem { get; set; }

        public bool IsTranslation { get; set; }

        public string Language { get; set; }

        public List<string> TranslatedLanguages { get; set; }

        public bool IsHomePage { get; set; }

        public bool IsTemplate { get; set; }

    }
}
