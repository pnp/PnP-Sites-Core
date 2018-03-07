using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// Static methods to modify URL paths.
    /// </summary>
    public static class UrlUtility
    {
        const char PATH_DELIMITER = '/';

#if !ONPREMISES
        const string INVALID_CHARS_REGEX = @"[\\#%*/:<>?+|\""]";
        const string REGEX_INVALID_FILEFOLDER_NAME_CHARS = @"[""#%*:<>?/\|\t\r\n]";
#else
        const string INVALID_CHARS_REGEX = @"[\\~#%&*{}/:<>?+|\""]";
        const string REGEX_INVALID_FILEFOLDER_NAME_CHARS = @"[~#%&*{}\:<>?/|""\t\r\n]";
#endif
        const string IIS_MAPPED_PATHS_REGEX = @"/(_layouts|_admin|_app_bin|_controltemplates|_login|_vti_bin|_vti_pvt|_windows|_wpresources)/";

        #region [ Combine ]
        /// <summary>
        /// Combines a path and a relative path.
        /// </summary>
        /// <param name="path">A SharePoint URL</param>
        /// <param name="relativePaths">SharePoint relative URLs</param>
        /// <returns>Returns comibed path with a relative paths</returns>
        public static string Combine(string path, params string[] relativePaths) {
            string pathBuilder = path ?? string.Empty;

            if (relativePaths == null)
                return pathBuilder;

            foreach (string relPath in relativePaths) {
                pathBuilder = Combine(pathBuilder, relPath);
            }
            return pathBuilder;
        }
        /// <summary>
        /// Combines a path and a relative path.
        /// </summary>
        /// <param name="path">A SharePoint URL</param>
        /// <param name="relative">SharePoint relative URL</param>
        /// <returns>Returns comibed path with a relative path</returns>
        public static string Combine(string path, string relative) 
        {
            if(relative == null)
                relative = string.Empty;
            
            if(path == null)
                path = string.Empty;

            if(relative.Length == 0 && path.Length == 0)
                return string.Empty;

            if(relative.Length == 0)
                return path;

            if(path.Length == 0)
                return relative;

            path = path.Replace('\\', PATH_DELIMITER);
            relative = relative.Replace('\\', PATH_DELIMITER);

            return path.TrimEnd(PATH_DELIMITER) + PATH_DELIMITER + relative.TrimStart(PATH_DELIMITER);
        }
        #endregion

        #region [ AppendQueryString ]
        /// <summary>
        /// Adds query string parameters to the end of a querystring and guarantees the proper concatenation with <b>?</b> and <b>&amp;.</b>
        /// </summary>
        /// <param name="path">A SharePoint URL</param>
        /// <param name="queryString">Query string value that need to append to the URL</param>
        /// <returns>Returns URL along with appended query string</returns>
        public static string AppendQueryString(string path, string queryString)
        {
            string url = path;

            if (queryString != null && queryString.Length > 0)
            {
                char startChar = (path.IndexOf("?") > 0) ? '&' : '?';
                url = string.Concat(path, startChar, queryString.TrimStart('?'));
            }
            return url;
        }
        #endregion

        #region [ RelativeUrl ]
        /// <summary>
        /// Returns realtive URL of given URL
        /// </summary>
        /// <param name="urlToProcess">SharePoint URL to process</param>
        /// <returns>Returns realtive URL of given URL</returns>
        public static string MakeRelativeUrl(string urlToProcess) {
            Uri uri = new Uri(urlToProcess);
            return uri.AbsolutePath;
        }

        /// <summary>
        /// Ensures that there is a trailing slash at the end of the URL
        /// </summary>
        /// <param name="urlToProcess"></param>
        /// <returns></returns>
        public static string EnsureTrailingSlash(string urlToProcess) 
        {
            if (!urlToProcess.EndsWith("/"))
            {
                return urlToProcess + "/";
            }

            return urlToProcess;
        }
        #endregion

        /// <summary>
        /// Checks if URL contains invalid characters or not
        /// </summary>
        /// <param name="content">Url value</param>
        /// <returns>Returns true if URL contains invalid characters. Otherwise returns false.</returns>
        public static bool ContainsInvalidUrlChars(this string content)
        {
	        return Regex.IsMatch(content, INVALID_CHARS_REGEX);
        }

        /// <summary>
        /// Checks if file or folder contains invalid characters or not
        /// </summary>
        /// <param name="content">File or folder name to check</param>
        /// <returns>True if contains invalid chars, false otherwise</returns>
        public static bool ContainsInvalidFileFolderChars(this string content)
        {
            return Regex.IsMatch(content, REGEX_INVALID_FILEFOLDER_NAME_CHARS);
        }

        /// <summary>
        /// Removes invalid characters
        /// </summary>
        /// <param name="content">Url value</param>
        /// <returns>Returns URL without invalid characters</returns>
        public static string StripInvalidUrlChars(this string content)
        {
            return ReplaceInvalidUrlChars(content, "");
        }
        /// <summary>
        /// Replaces invalid charcters with other characters
        /// </summary>
        /// <param name="content">Url value</param>
        /// <param name="replacer">string need to replace with invalid characters</param>
        /// <returns>Returns replaced invalid charcters from URL</returns>
        public static string ReplaceInvalidUrlChars(this string content, string replacer)
        {
	    return new Regex(INVALID_CHARS_REGEX).Replace(content, replacer);
        }

        /// <summary>
        /// Tells URL is virtual directory or not
        /// </summary>
        /// <param name="url">SharePoint URL</param>
        /// <returns>Returns true if URL is virtual directory. Otherwise returns false.</returns>
        public static bool IsIisVirtualDirectory(string url)
        {
            return Regex.IsMatch(url, IIS_MAPPED_PATHS_REGEX, RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Taken from Microsoft.SharePoint.Utilities.SPUtility
        /// </summary>
        /// <param name="strUrl"></param>
        /// <param name="strBaseUrl"></param>
        /// <returns></returns>
        internal static string ConvertToServiceRelUrl(string strUrl, string strBaseUrl)
        {
            if (((strBaseUrl == null) || !StsStartsWith(strBaseUrl, "/")) || ((strUrl == null) || !StsStartsWith(strUrl, "/")))
            {
                throw new ArgumentException();
            }
            if ((strUrl.Length > 1) && (strUrl[strUrl.Length - 1] == '/'))
            {
                strUrl = strUrl.Substring(0, strUrl.Length - 1);
            }
            if ((strBaseUrl.Length > 1) && (strBaseUrl[strBaseUrl.Length - 1] == '/'))
            {
                strBaseUrl = strBaseUrl.Substring(0, strBaseUrl.Length - 1);
            }
            if (!StsStartsWith(strUrl, strBaseUrl))
            {
                throw new ArgumentException();
            }
            if (strBaseUrl == "/")
            {
                return strUrl.Substring(1);
            }
            if (strUrl.Length == strBaseUrl.Length)
            {
                return "";
            }
            return strUrl.Substring(strBaseUrl.Length + 1);
        }

        /// <summary>
        /// Taken from Microsoft.SharePoint.Utilities.SPUtility
        /// </summary>
        /// <param name="strMain"></param>
        /// <param name="strBegining"></param>
        /// <returns></returns>
        internal static bool StsStartsWith(string strMain, string strBegining)
        {
            return CultureInfo.InvariantCulture.CompareInfo.IsPrefix(strMain, strBegining, CompareOptions.IgnoreCase);
        }
    }
}
