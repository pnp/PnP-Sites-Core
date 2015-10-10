using Microsoft.SharePoint.Client.Search.Administration;
using Microsoft.SharePoint.Client.Search.Portability;
using System;
using System.Text;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for Search extension methods
    /// </summary>
    public static partial class SearchExtensions
    {
        /// <summary>
        /// Returns the current search configuration for the specified object level
        /// </summary>
        /// <param name="context"></param>
        /// <param name="searchSettingsObjectLevel"></param>
        /// <returns></returns>
        public static string GetSearchConfiguration(this ClientContext context, SearchObjectLevel searchSettingsObjectLevel)
        {
            return GetSearchConfigurationImplementation(context, searchSettingsObjectLevel);
        }
    }
}
