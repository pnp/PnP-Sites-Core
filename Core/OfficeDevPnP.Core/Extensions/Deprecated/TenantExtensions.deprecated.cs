using System;
using System.ComponentModel;
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for tenant extension methods
    /// </summary>
    public static partial class TenantExtensions
    {
#if !ONPREMISES
        /// <summary>
        /// Adds a package to the tenants app catalog and by default deploys it if the package is a client side package (sppkg)
        /// </summary>
        /// <param name="tenant">Tenant to operate against</param>
        /// <param name="spPkgName">Name of the package to upload (e.g. demo.sppkg) </param>
        /// <param name="spPkgPath">Path on the filesystem where this package is stored</param>
        /// <param name="autoDeploy">Automatically deploy the package, only applies to client side packages (sppkg)</param>
        /// <param name="overwrite">Overwrite the package if it was already listed in the app catalog</param>
        /// <returns>The ListItem of the added package row</returns>
        [Obsolete("Please use the DeployApplicationPackageToAppCatalog extension method on the Web class. This method will be removed in the October 2017 release.")]
        public static ListItem DeployApplicationPackageToAppCatalog(this Tenant tenant, string spPkgName, string spPkgPath, bool autoDeploy = true, bool overwrite = true)
        {
            var appCatalogSite = tenant.GetAppCatalog();
            if (appCatalogSite == null)
            {
                throw new ArgumentException("No app catalog site found, please ensure the site exists or specify the site as parameter. Note that the app catalog site is retrieved via search, so take in account the indexing time.");
            }

            return DeployApplicationPackageToAppCatalogImplementation(tenant, appCatalogSite.ToString(), spPkgName, spPkgPath, autoDeploy, false, overwrite);
        }

        /// <summary>
        /// Adds a package to the tenants app catalog and by default deploys it if the package is a client side package (sppkg)
        /// </summary>
        /// <param name="tenant">Tenant to operate against</param>
        /// <param name="spPkgName">Name of the package to upload (e.g. demo.sppkg) </param>
        /// <param name="spPkgPath">Path on the filesystem where this package is stored</param>
        /// <param name="autoDeploy">Automatically deploy the package, only applies to client side packages (sppkg)</param>
        /// <param name="skipFeatureDeployment">Skip the feature deployment step, allows for a one-time central deployment of your solution</param>
        /// <param name="overwrite">Overwrite the package if it was already listed in the app catalog</param>
        /// <returns>The ListItem of the added package row</returns>
        [Obsolete("Please use the DeployApplicationPackageToAppCatalog extension method on the Web class. This method will be removed in the October 2017 release.")]
        public static ListItem DeployApplicationPackageToAppCatalog(this Tenant tenant, string spPkgName, string spPkgPath, bool autoDeploy = true, bool skipFeatureDeployment = true, bool overwrite = true)
        {
            var appCatalogSite = tenant.GetAppCatalog();
            if (appCatalogSite == null)
            {
                throw new ArgumentException("No app catalog site found, please ensure the site exists or specify the site as parameter. Note that the app catalog site is retrieved via search, so take in account the indexing time.");
            }

            return DeployApplicationPackageToAppCatalogImplementation(tenant, appCatalogSite.ToString(), spPkgName, spPkgPath, autoDeploy, skipFeatureDeployment, overwrite);
        }

        /// <summary>
        /// Adds a package to the tenants app catalog and by default deploys it if the package is a client side package (sppkg)
        /// </summary>
        /// <param name="tenant">Tenant to operate against</param>
        /// <param name="appCatalogSiteUrl">Full URL to the tenant admin site (e.g. https://contoso.sharepoint.com/sites/apps) </param>
        /// <param name="spPkgName">Name of the package to upload (e.g. demo.sppkg) </param>
        /// <param name="spPkgPath">Path on the filesystem where this package is stored</param>
        /// <param name="autoDeploy">Automatically deploy the package, only applies to client side packages (sppkg)</param>
        /// <param name="overwrite">Overwrite the package if it was already listed in the app catalog</param>
        /// <returns>The ListItem of the added package row</returns>
        [Obsolete("Please use the DeployApplicationPackageToAppCatalog overloads on the Web class that don't require you to specify the appCatalogSiteUrl parameter. This method will be removed in the October 2017 release.")]
        public static ListItem DeployApplicationPackageToAppCatalog(this Tenant tenant, string appCatalogSiteUrl, string spPkgName, string spPkgPath, bool autoDeploy = true, bool overwrite = true)
        {
            return DeployApplicationPackageToAppCatalogImplementation(tenant, appCatalogSiteUrl, spPkgName, spPkgPath, autoDeploy, false, overwrite);
        }

        private static ListItem DeployApplicationPackageToAppCatalogImplementation(this Tenant tenant, string appCatalogSiteUrl, string spPkgName, string spPkgPath, bool autoDeploy, bool skipFeatureDeployment, bool overwrite)
        {
            if (String.IsNullOrEmpty(appCatalogSiteUrl))
            {
                throw new ArgumentException("Please specify a app catalog site URL");
            }

            Uri catalogUri;
            if (!Uri.TryCreate(appCatalogSiteUrl, UriKind.Absolute, out catalogUri))
            {
                throw new ArgumentException("Please specify a valid app catalog site URL");
            }

            if (String.IsNullOrEmpty(spPkgName))
            {
                throw new ArgumentException("Please specify a package name");
            }

            if (String.IsNullOrEmpty(spPkgPath))
            {
                throw new ArgumentException("Please specify a package path");
            }

            using (var appCatalogContext = tenant.Context.Clone(catalogUri))
            {
                List catalog = appCatalogContext.Web.GetListByUrl("appcatalog");
                if (catalog == null)
                {
                    throw new Exception($"No app catalog found...did you provide a valid app catalog site?");
                }

                Folder rootFolder = catalog.RootFolder;

                // Upload package
                var sppkgFile = rootFolder.UploadFile(spPkgName, System.IO.Path.Combine(spPkgPath, spPkgName), overwrite);
                if (sppkgFile == null)
                {
                    throw new Exception($"Upload of {spPkgName} failed");
                }

                if ((autoDeploy || skipFeatureDeployment) &&
                    System.IO.Path.GetExtension(spPkgName).ToLower() == ".sppkg")
                {
                    // Trigger "deployment" by setting the IsClientSideSolutionDeployed bool to true which triggers 
                    // an event receiver that will process the sppkg file and update the client side componenent manifest list
                    sppkgFile.ListItemAllFields["IsClientSideSolutionDeployed"] = autoDeploy;
                    // deal with "upgrading" solutions
                    sppkgFile.ListItemAllFields["IsClientSideSolutionCurrentVersionDeployed"] = autoDeploy;
                    // Allow for a central deployment of the solution, no need to install the solution in the individual site collections.
                    // Only works when the solution is not using feature framework to "configure" the site upon solution installation
                    sppkgFile.ListItemAllFields["SkipFeatureDeployment"] = skipFeatureDeployment;
                    sppkgFile.ListItemAllFields.Update();
                }

                appCatalogContext.Load(sppkgFile.ListItemAllFields);
                appCatalogContext.ExecuteQueryRetry();

                return sppkgFile.ListItemAllFields;
            }
        }
#endif
    }
}
