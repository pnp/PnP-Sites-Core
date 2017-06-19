using System;
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
#else
        const string INVALID_CHARS_REGEX = @"[\\~#%&*{}/:<>?+|\""]";
#endif
        const string IIS_MAPPED_PATHS_REGEX = @"/(_layouts|_admin|_app_bin|_controltemplates|_login|_vti_bin|_vti_pvt|_windows|_wpresources)/";

        #region [ Combine ]
        /// <summary>
        /// Combines a path and a relative path.
        /// </summary>
        /// <param name="path">A SharePoint url</param>
        /// <param name="relativePaths">SharePoint relative urls</param>
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
        /// <param name="path">A SharePoint url</param>
        /// <param name="relative">SharePoint relative url</param>
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
        /// <param name="path">A SharePoint url</param>
        /// <param name="queryString">Query string value that need to append to the url</param>
        /// <returns>Returns url along with appended query string</returns>
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
        /// Returns realtive url of given url
        /// </summary>
        /// <param name="urlToProcess">SharePoint url to process</param>
        /// <returns>Returns realtive url of given url</returns>
        public static string MakeRelativeUrl(string urlToProcess) {
            Uri uri = new Uri(urlToProcess);
            return uri.AbsolutePath;
        }

        /// <summary>
        /// Ensures that there is a trailing slash at the end of the url
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
        /// Checks url contians invalid characters or not
        /// </summary>
        /// <param name="content">url value</param>
        /// <returns>Returns true if url contains invalid characters. Otherwise returns false.</returns>
        public static bool ContainsInvalidUrlChars(this string content)
        {
	    return Regex.IsMatch(content, INVALID_CHARS_REGEX);
        }

        /// <summary>
        /// Removes invalid characters
        /// </summary>
        /// <param name="content">url value</param>
        /// <returns>Returns url without invalid characters</returns>
        public static string StripInvalidUrlChars(this string content)
        {
            return ReplaceInvalidUrlChars(content, "");
        }
        /// <summary>
        /// Replaces invalid charcters with other characters
        /// </summary>
        /// <param name="content">Url value</param>
        /// <param name="replacer">string need to replace with invalid characters</param>
        /// <returns>Returns replaced invalid charcters from url</returns>
        public static string ReplaceInvalidUrlChars(this string content, string replacer)
        {
	    return new Regex(INVALID_CHARS_REGEX).Replace(content, replacer);
        }

        /// <summary>
        /// Tells url is virtual directory or not
        /// </summary>
        /// <param name="url">SharePoint url</param>
        /// <returns>Returns true if url is virtual directory. Otherwise returns false.</returns>
        public static bool IsIisVirtualDirectory(string url)
        {
            return Regex.IsMatch(url, IIS_MAPPED_PATHS_REGEX, RegexOptions.IgnoreCase);
        }

    }
}
