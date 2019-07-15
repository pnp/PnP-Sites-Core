using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Async;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.ALM
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Allows Application Lifecycle Management for Apps
    /// </summary>
    public class AppManager
    {
        private ClientContext _context;
        private enum AppManagerAction
        {
            Install,
            Retract,
            Remove,
            Deploy,
            Upgrade,
            Uninstall
        }

        public AppManager(ClientContext context)
        {
            //_context = context ?? throw new ArgumentException(nameof(context));
            if (context == null)
            {
                throw new ArgumentException(nameof(context));
            }
            else
            {
                _context = context;
            }
        }

        /// <summary>
        /// Uploads a file to the Tenant App Catalog
        /// </summary>
        /// <param name="file">A byte array containing the file</param>
        /// <param name="filename">The filename (e.g. myapp.sppkg) of the file to upload</param>
        /// <param name="overwrite">If true will overwrite an existing entry</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <param name="timeoutSeconds">If specified will set the timeout on the request. Defaults to 200 seconds.</param>
        /// <returns></returns>
        public AppMetadata Add(byte[] file, string filename, bool overwrite = false, AppCatalogScope scope = AppCatalogScope.Tenant, int timeoutSeconds = 200)
        {
            return Task.Run(() => AddAsync(file, filename, overwrite, scope, timeoutSeconds)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Uploads an app file to the Tenant App Catalog
        /// </summary>
        /// <param name="path"></param>
        /// <param name="overwrite">If true will overwrite an existing entry</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <param name="timeoutSeconds">If specified will set the timeout on the request. Defaults to 200 seconds.</param>
        /// <returns></returns>
        public AppMetadata Add(string path, bool overwrite = false, AppCatalogScope scope = AppCatalogScope.Tenant, int timeoutSeconds = 200)
        {
            return Task.Run(() => AddAsync(path, overwrite, scope, timeoutSeconds)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Uploads a file to the Tenant App Catalog
        /// </summary>
        /// <param name="file">A byte array containing the file</param>
        /// <param name="filename">The filename (e.g. myapp.sppkg) of the file to upload</param>
        /// <param name="overwrite">If true will overwrite an existing entry</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <param name="timeoutSeconds">If specified will set the timeout on the request. Defaults to 200 seconds.</param>
        /// <returns></returns>
        public async Task<AppMetadata> AddAsync(byte[] file, string filename, bool overwrite = false, AppCatalogScope scope = AppCatalogScope.Tenant, int timeoutSeconds = 200)
        {
            if (file == null && file.Length == 0)
            {
                throw new ArgumentException(nameof(file));
            }
            if (string.IsNullOrEmpty(filename))
            {
                throw new ArgumentException(nameof(filename));
            }

            await new SynchronizationContextRemover();

            return await BaseAddRequest(file, filename, overwrite, timeoutSeconds, scope);
        }

        /// <summary>
        /// Uploads an app file to the Tenant App Catalog
        /// </summary>
        /// <param name="path"></param>
        /// <param name="overwrite">If true will overwrite an existing entry</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <param name="timeoutSeconds">If specified will set the timeout on the request. Defaults to 200 seconds.</param>
        /// <returns></returns>
        public async Task<AppMetadata> AddAsync(string path, bool overwrite = false, AppCatalogScope scope = AppCatalogScope.Tenant, int timeoutSeconds = 200)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException(nameof(path));
            }

            if (!System.IO.File.Exists(path))
            {
                throw new IOException("File does not exist");
            }

            var bytes = System.IO.File.ReadAllBytes(path);
            var fileInfo = new FileInfo(path);

            await new SynchronizationContextRemover();

            return await BaseAddRequest(bytes, fileInfo.Name, overwrite, timeoutSeconds, scope);
        }

        /// <summary>
        /// Installs an available app from the app catalog in a site.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to install</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Install(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => InstallAsync(appMetadata, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Installs an available app from the app catalog in a site.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to install</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> InstallAsync(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }

            await new SynchronizationContextRemover();

            return await InstallAsync(appMetadata.Id, scope);
        }

        /// <summary>
        /// Installs an available app from the app catalog in a site.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Install(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => InstallAsync(id, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Installs an available app from the app catalog in a site.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> InstallAsync(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest(id, AppManagerAction.Install, false, null, scope);
        }

        /// <summary>
        /// Uninstalls an app from a site.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to uninstall.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Uninstall(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => UninstallAsync(appMetadata, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Uninstalls an app from a site.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to uninstall.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> UninstallAsync(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }

            await new SynchronizationContextRemover();

            return await UninstallAsync(appMetadata.Id, scope);
        }

        /// <summary>
        /// Uninstalls an app from a site.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Uninstall(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => UninstallAsync(id, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Uninstalls an app from a site.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> UninstallAsync(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest(id, AppManagerAction.Uninstall, false, null, scope);
        }

        /// <summary>
        /// Upgrades an app in a site
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to upgrade.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Upgrade(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => UpgradeAsync(appMetadata, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Upgrades an app in a site
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to upgrade.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> UpgradeAsync(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }

            await new SynchronizationContextRemover();

            return await UpgradeAsync(appMetadata.Id, scope);
        }

        /// <summary>
        /// Upgrades an app in a site
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Upgrade(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => UpgradeAsync(id, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Upgrades an app in a site
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> UpgradeAsync(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest(id, AppManagerAction.Upgrade, false, null, scope);
        }

        /// <summary>
        /// Deploys/trusts an app in the app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to deploy.</param>
        /// <param name="skipFeatureDeployment">If set to true will skip the feature deployed for tenant scoped apps.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Deploy(AppMetadata appMetadata, bool skipFeatureDeployment = true, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => DeployAsync(appMetadata, skipFeatureDeployment, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Deploys/trusts an app in the app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to deploy.</param>
        /// <param name="skipFeatureDeployment">If set to true will skip the feature deployed for tenant scoped apps.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> DeployAsync(AppMetadata appMetadata, bool skipFeatureDeployment = true, AppCatalogScope scope = AppCatalogScope.Tenant)
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

            await new SynchronizationContextRemover();

            return await BaseRequest(appMetadata.Id, AppManagerAction.Deploy, true, postObj, scope);
        }

        /// <summary>
        /// Deploys/trusts an app in the app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="skipFeatureDeployment">If set to true will skip the feature deployed for tenant scoped apps.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Deploy(Guid id, bool skipFeatureDeployment = true, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => DeployAsync(id, skipFeatureDeployment, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Deploys/trusts an app in the app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="skipFeatureDeployment">If set to true will skip the feature deployed for tenant scoped apps.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> DeployAsync(Guid id, bool skipFeatureDeployment = true, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }
            var postObj = new Dictionary<string, object>
            {
                { "skipFeatureDeployment", skipFeatureDeployment }
            };

            await new SynchronizationContextRemover();

            return await BaseRequest(id, AppManagerAction.Deploy, true, postObj, scope);
        }

        /// <summary>
        /// Retracts an app in the app catalog. Notice that this will not remove the app from the app catalog.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to retract.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Retract(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => RetractAsync(appMetadata, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Retracts an app in the app catalog. Notice that this will not remove the app from the app catalog.
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to retract.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> RetractAsync(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest(appMetadata.Id, AppManagerAction.Retract, true, null, scope);
        }

        /// <summary>
        /// Retracts an app in the app catalog. Notice that this will not remove the app from the app catalog.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Retract(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => RetractAsync(id, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Retracts an app in the app catalog. Notice that this will not remove the app from the app catalog.
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> RetractAsync(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest(id, AppManagerAction.Retract, true, null, scope);
        }

        /// <summary>
        /// Removes an app from the app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to remove.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Remove(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => RemoveAsync(appMetadata, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Removes an app from the app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to remove.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> RemoveAsync(AppMetadata appMetadata, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (appMetadata == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }
            if (appMetadata.Id == Guid.Empty)
            {
                throw new ArgumentException(nameof(appMetadata.Id));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest(appMetadata.Id, AppManagerAction.Remove, true, null, scope);
        }

        /// <summary>
        /// Removes an app from the app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public bool Remove(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => RemoveAsync(id, scope)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Removes an app from the app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<bool> RemoveAsync(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest(id, AppManagerAction.Remove, true, null, scope);
        }

        /// <summary>
        /// Synchronize an app from the tenant app catalog with the teams app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listen in the app catalog</param>
        /// <returns></returns>
        public async Task<bool> SyncToTeamsAsync(Guid id)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }

            await new SynchronizationContextRemover();

            return await SyncToTeamsImplementation(id);
        }

        /// <summary>
        /// Synchronize an app from the tenant app catalog with the teams app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to remove.</param>
        /// <returns></returns>
        public async Task<bool> SyncToTeamsAsync(AppMetadata appMetadata)
        {
            if (appMetadata == null || appMetadata.Id == null)
            {
                throw new ArgumentException(nameof(appMetadata));
            }

            await new SynchronizationContextRemover();

            return await SyncToTeamsImplementation(appMetadata.Id);
        }

        /// <summary>
        /// Synchronize an app from the tenant app catalog with the teams app catalog
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listen in the app catalog</param>
        /// <returns></returns>
        public bool SyncToTeams(Guid id)
        {
            return Task.Run(() => SyncToTeamsAsync(id)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Synchronize an app from the tenant app catalog with the teams app catalog
        /// </summary>
        /// <param name="appMetadata">The app metadata object of the app to remove.</param>
        /// <returns></returns>
        public bool SyncToTeams(AppMetadata appMetadata)
        {
            return Task.Run(() => SyncToTeamsAsync(appMetadata)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns all available apps.
        /// </summary>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public List<AppMetadata> GetAvailable(AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => BaseGetAvailableAsync(scope, Guid.Empty)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns all available apps.
        /// </summary>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<List<AppMetadata>> GetAvailableAsync(AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            await new SynchronizationContextRemover();

            return await BaseGetAvailableAsync(scope, Guid.Empty);
        }

        /// <summary>
        /// Returns an available app
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public AppMetadata GetAvailable(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => BaseGetAvailableAsync(scope, id)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns an available app
        /// </summary>
        /// <param name="id">The id of the app</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<AppMetadata> GetAvailableAsync(Guid id, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            await new SynchronizationContextRemover();

            return await BaseGetAvailableAsync(scope, id);
        }

        /// <summary>
        /// Returns an available app
        /// </summary>
        /// <param name="title">The title of the app.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public AppMetadata GetAvailable(string title, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            return Task.Run(() => BaseGetAvailableAsync(scope, Guid.Empty, title)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns an avialable app
        /// </summary>
        /// <param name="title">The title of the app.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        public async Task<AppMetadata> GetAvailableAsync(string title, AppCatalogScope scope = AppCatalogScope.Tenant)
        {
            await new SynchronizationContextRemover();

            return await BaseGetAvailableAsync(scope, Guid.Empty, title);
        }

        #region Private Methods

        /// <summary>
        /// Returns an available app
        /// </summary>
        /// <param name="id">The unique id of the app. Notice that this is not the product id as listed in the app catalog.</param>
        /// <param name="title">The title of the app.</param>
        /// <param name="scope">Specifies the app catalog to work with. Defaults to Tenant</param>
        /// <returns></returns>
        private async Task<dynamic> BaseGetAvailableAsync(AppCatalogScope scope, Guid id = default(Guid), string title = "")
        {
            dynamic addins = null;

            var accessToken = _context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                _context.Web.EnsureProperty(w => w.Url);

                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(_context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {

                    string requestUrl = $"{_context.Web.Url}/_api/web/{(scope == AppCatalogScope.Tenant ? "tenant" : "sitecollection")}appcatalog/AvailableApps";
                    if (Guid.Empty != id)
                    {
                        requestUrl = $"{_context.Web.Url}/_api/web/{(scope == AppCatalogScope.Tenant ? "tenant" : "sitecollection")}appcatalog/AvailableApps/GetById('{id}')";
                    }
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (_context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }
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
                                if (Guid.Empty == id && string.IsNullOrEmpty(title))
                                {
                                    var responseJson = JObject.Parse(responseString);
                                    var returnedAddins = responseJson["d"]["results"] as JArray;

                                    addins = JsonConvert.DeserializeObject<List<AppMetadata>>(returnedAddins.ToString());
                                }
                                else if (!String.IsNullOrEmpty(title))
                                {
                                    var responseJson = JObject.Parse(responseString);
                                    var returnedAddins = responseJson["d"]["results"] as JArray;

                                    var listAddins = JsonConvert.DeserializeObject<List<AppMetadata>>(returnedAddins.ToString());
                                    addins = listAddins.Where(a => a.Title == title).FirstOrDefault();
                                }
                                else
                                {
                                    var responseJson = JObject.Parse(responseString);
                                    var returnedAddins = responseJson["d"];
                                    addins = JsonConvert.DeserializeObject<AppMetadata>(returnedAddins.ToString());
                                }

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


        private async Task<bool> BaseRequest(Guid id, AppManagerAction action, bool switchToAppCatalogContext, Dictionary<string, object> postObject, AppCatalogScope scope, int timeoutSeconds = 200)
        {
            var context = _context;
            if (switchToAppCatalogContext == true && scope == AppCatalogScope.Tenant)
            {
                // switch context to appcatalog
                var appcatalogUri = _context.Web.GetAppCatalog();
                context = context.Clone(appcatalogUri);
            }
            var returnValue = false;
            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }


                using (var httpClient = new PnPHttpProvider(handler))
                {
                    httpClient.Timeout = new TimeSpan(0, 0, timeoutSeconds);

                    var method = action.ToString();
                    var requestUrl = $"{context.Web.Url}/_api/web/{(scope == AppCatalogScope.Tenant ? "tenant" : "sitecollection")}appcatalog/AvailableApps/GetByID('{id}')/{method}";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=nometadata");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());

                    if (postObject != null)
                    {
                        var jsonBody = JsonConvert.SerializeObject(postObject);
                        var requestBody = new StringContent(jsonBody);
                        MediaTypeHeaderValue sharePointJsonMediaType;
                        MediaTypeHeaderValue.TryParse("application/json;odata=nometadata;charset=utf-8", out sharePointJsonMediaType);
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

        private async Task<bool> SyncToTeamsImplementation(Guid appId)
        {
            var context = _context;

            // switch context to appcatalog
            var appcatalogUri = _context.Web.GetAppCatalog();
            context = context.Clone(appcatalogUri);

            var returnValue = false;
            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // find the app by id

                var list = context.Web.GetListByUrl("appcatalog");
                var query = new CamlQuery();
                query.ViewXml = $"<View><Query><Where><Contains><FieldRef Name='UniqueId'/><Value Type='Text'>{appId}</Value></Contains></Where></Query></View>";
                var items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQueryRetry();
                
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }
                if (items.Count > 0)
                {
                    using (var httpClient = new PnPHttpProvider(handler))
                    {
                        var requestUrl = $"{context.Web.Url}/_api/web/tenantappcatalog/SyncSolutionToTeams(id={items[0].Id})";

                        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                        request.Headers.Add("accept", "application/json;odata=nometadata");
                        if (!string.IsNullOrEmpty(accessToken))
                        {
                            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                        }
                        else
                        {
                            if (context.Credentials is NetworkCredential networkCredential)
                            {
                                handler.Credentials = networkCredential;
                            }
                        }
                        request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());

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
            }
            return await Task.Run(() => returnValue);
        }

        private async Task<AppMetadata> BaseAddRequest(byte[] file, string filename, bool overwrite, int timeoutSeconds, AppCatalogScope scope)
        {
            AppMetadata returnValue = null;

            var context = _context;
            if (scope == AppCatalogScope.Tenant)
            {
                // switch context to appcatalog
                var appcatalogUri = _context.Web.GetAppCatalog();
                context = context.Clone(appcatalogUri);
            }

            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {

                    string requestUrl = $"{context.Web.Url}/_api/web/{(scope == AppCatalogScope.Tenant ? "tenant" : "sitecollection")}appcatalog/Add(overwrite={(overwrite.ToString().ToLower())}, url='{filename}')";

                    var requestDigest = await context.GetRequestDigest();
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }
                    request.Headers.Add("X-RequestDigest", requestDigest);
                    request.Headers.Add("binaryStringRequestBody", "true");
                    request.Content = new ByteArrayContent(file);
                    httpClient.Timeout = new TimeSpan(0, 0, timeoutSeconds);
                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            var responseJson = JObject.Parse(responseString);
                            var id = responseJson["d"]["UniqueId"].ToString();

                            var metadataRequestUrl = $"{context.Web.Url}/_api/web/{(scope == AppCatalogScope.Tenant ? "tenant" : "sitecollection")}appcatalog/AvailableApps/GetById('{id}')";

                            HttpRequestMessage metadataRequest = new HttpRequestMessage(HttpMethod.Post, metadataRequestUrl);
                            metadataRequest.Headers.Add("accept", "application/json;odata=verbose");
                            if (!string.IsNullOrEmpty(accessToken))
                            {
                                metadataRequest.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                            }
                            metadataRequest.Headers.Add("X-RequestDigest", requestDigest);

                            // Perform actual post operation
                            HttpResponseMessage metadataResponse = await httpClient.SendAsync(metadataRequest, new System.Threading.CancellationToken());

                            if (metadataResponse.IsSuccessStatusCode)
                            {
                                // If value empty, URL is taken
                                var metadataResponseString = await metadataResponse.Content.ReadAsStringAsync();
                                if (metadataResponseString != null)
                                {
                                    var metadataResponseJson = JObject.Parse(metadataResponseString);
                                    var returnedAddins = metadataResponseJson["d"];
                                    returnValue = JsonConvert.DeserializeObject<AppMetadata>(returnedAddins.ToString());
                                }
                            }
                            else
                            {
                                // Something went wrong...
                                throw new Exception(await metadataResponse.Content.ReadAsStringAsync());
                            }
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
#endif
}
