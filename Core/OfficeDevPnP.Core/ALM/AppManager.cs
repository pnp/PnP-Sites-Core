using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.ALM
{
    /// <summary>
    /// Allows Application Lifecycle Management for Apps
    /// </summary>
    public class AppManager
    {
        private ClientContext _context;
        public AppManager(ClientContext context)
        {
            _context = context ?? throw new ArgumentException(nameof(context));
        }

        /// <summary>
        /// Uploads a file to the Tenant App Catalog
        /// </summary>
        /// <param name="file">A byte array containing the file</param>
        /// <param name="filename">The filename (e.g. myapp.sppkg) of the file to upload</param>
        /// <returns></returns>
        public async Task<bool> Add(byte[] file, string filename)
        {
            if (file == null && file.Length == 0)
            {
                throw new ArgumentException(nameof(file));
            }
            if (string.IsNullOrEmpty(filename))
            {
                throw new ArgumentException(nameof(filename));
            }
            return await BaseAddRequest(file, filename, true);
        }

        /// <summary>
        /// Uploads an app file to the Tenant App Catalog
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public async Task<bool> Add(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException(nameof(path));
            }

            if (System.IO.File.Exists(path))
            {
                throw new IOException("File does not exist");
            }

            var bytes = System.IO.File.ReadAllBytes(path);
            var fileInfo = new FileInfo(path);
            return await BaseAddRequest(bytes, fileInfo.Name, true);
        }



        /// <summary>
        /// Installs an available app from the app catalog in a site.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to install</param>
        /// <returns></returns>
        public async Task<bool> InstallAsync(AppMetadata appMetadata)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }
            return await InstallAsync(appMetadata.Id);
        }

        /// <summary>
        /// Installs an available app from the app catalog in a site.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <returns></returns>
        public async Task<bool> InstallAsync(Guid id)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }
            return await BaseRequest(id, "Install");
        }

        /// <summary>
        /// Uninstalls an app from a site.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to uninstall.</param>
        /// <returns></returns>
        public async Task<bool> UninstallAsync(AppMetadata appMetadata)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }
            return await UninstallAsync(appMetadata.Id);
        }

        /// <summary>
        /// Uninstalls an app from a site.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <returns></returns>
        public async Task<bool> UninstallAsync(Guid id)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }
            return await BaseRequest(id, "Uninstall");
        }

        /// <summary>
        /// Upgrades an app in a site
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to upgrade.</param>
        /// <returns></returns>
        public async Task<bool> UpgradeAsync(AppMetadata appMetadata)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }
            return await UpgradeAsync(appMetadata.Id);
        }

        /// <summary>
        /// Upgrades an app in a site
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <returns></returns>
        public async Task<bool> UpgradeAsync(Guid id)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }
            return await BaseRequest(id, "Upgrade");
        }

        /// <summary>
        /// Deploys/trusts an app in the app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to deploy.</param>
        /// <param name="skipFeatureDeployment">If set to true will skip the feature deployed for tenant scoped apps.</param>
        /// <returns></returns>
        public async Task<bool> DeployAsync(AppMetadata appMetadata, bool skipFeatureDeployment = true)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }
            var postObj = new Dictionary<string, object>
            {
                { "skipFeatureDeployment", skipFeatureDeployment }
            };
            return await BaseRequest(appMetadata.Id, "Deploy", true, postObj);
        }

        /// <summary>
        /// Deploys/trusts an app in the app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="skipFeatureDeployment">If set to true will skip the feature deployed for tenant scoped apps.</param>
        /// <returns></returns>
        public async Task<bool> DeployAsync(Guid id, bool skipFeatureDeployment = true)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }
            var postObj = new Dictionary<string, object>
            {
                { "skipFeatureDeployment", skipFeatureDeployment }
            };
            return await BaseRequest(id, "Deploy", true, postObj);
        }

        /// <summary>
        /// Retracts an app in the app catalog. Notice that this will not remove the app from the app catalog.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to retract.</param>
        /// <returns></returns>
        public async Task<bool> RetractAsync(AppMetadata appMetadata)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }
            return await BaseRequest(appMetadata.Id, "Retract", true);
        }

        /// <summary>
        /// Retracts an app in the app catalog. Notice that this will not remove the app from the app catalog.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <returns></returns>
        public async Task<bool> RetractAsync(Guid id)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }
            return await BaseRequest(id, "Retract", true);
        }

        /// <summary>
        /// Removes an app from the app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to remove.</param>
        /// <returns></returns>
        public async Task<bool> RemoveAsync(AppMetadata appMetadata)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }
            return await BaseRequest(appMetadata.Id, "Remove", true);
        }

        /// <summary>
        /// Removes an app from the app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <returns></returns>
        public async Task<bool> RemoveAsync(Guid id)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }
            return await BaseRequest(id, "Remove", true);
        }

        public async Task<List<AppMetadata>> GetAvailableAddinsAsync()
        {
            List<AppMetadata> addins = new List<AppMetadata>();

            var accessToken = _context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    _context.Web.EnsureProperty(w => w.Url);
                    handler.Credentials = _context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(_context.Web.Url), (_context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(_context.Web.Url)));
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = String.Format("{0}/_api/web/tenantappcatalog/AvailableApps", _context.Web.Url);

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    MediaTypeHeaderValue sharePointJsonMediaType = null;
                    MediaTypeHeaderValue.TryParse("application/json;odata=verbose;charset=utf-8", out sharePointJsonMediaType);

                    request.Headers.Add("X-RequestDigest", await _context.GetRequestDigest());

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            try
                            {
                                var responseJson = JObject.Parse(responseString);
                                var returnedAddins = responseJson["d"]["results"] as JArray;

                                addins = JsonConvert.DeserializeObject<List<AppMetadata>>(returnedAddins.ToString());

                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return await Task.Run(() => addins);
        }

        #region Private Methods
        private async Task<bool> BaseRequest(Guid id, string method, bool appCatalog = false, Dictionary<string, object> postObject = null)
        {
            var context = _context;
            if (appCatalog == true)
            {
                // switch context to appcatalog
                var appcatalogUri = _context.Web.GetAppCatalog();
                context = context.Clone(appcatalogUri);
            }
            var returnValue = false;
            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    context.Web.EnsureProperty(w => w.Url);
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = $"{context.Web.Url}/_api/web/tenantappcatalog/AvailableApps/GetByID('{id}')/{method}";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=nometadata");
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());

                    if (postObject != null)
                    {
                        var jsonBody = JsonConvert.SerializeObject(postObject);
                        var requestBody = new StringContent(jsonBody);
                        MediaTypeHeaderValue.TryParse("application/json;odata=nometadata;charset=utf-8", out MediaTypeHeaderValue sharePointJsonMediaType);
                        requestBody.Headers.ContentType = sharePointJsonMediaType;
                        request.Content = requestBody;
                    }

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            try
                            {
                                var responseJson = JObject.Parse(responseString);
                                returnValue = true;
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return await Task.Run(() => returnValue);
        }

        private async Task<bool> BaseAddRequest(byte[] file, string filename, bool overwrite = false, bool appCatalog = true)
        {
            var context = _context;
            if (appCatalog == true)
            {
                // switch context to appcatalog
                var appcatalogUri = _context.Web.GetAppCatalog();
                context = context.Clone(appcatalogUri);
            }
            var returnValue = false;
            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    context.Web.EnsureProperty(w => w.Url);
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = $"{context.Web.Url}/_api/web/tenantappcatalog/Add(overwrite={(overwrite.ToString().ToLower())}, url='{filename}')";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    MediaTypeHeaderValue sharePointJsonMediaType = null;
                    MediaTypeHeaderValue.TryParse("application/json;odata=verbose;charset=utf-8", out sharePointJsonMediaType);
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());
                    request.Headers.Add("binaryStringRequestBody", "true");
                    request.Content = new ByteArrayContent(file);

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            try
                            {
                                var responseJson = JObject.Parse(responseString);
                                returnValue = true;
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return await Task.Run(() => returnValue);
        }
        #endregion
    }
}
