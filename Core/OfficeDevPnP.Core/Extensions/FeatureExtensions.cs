using System;
using System.Linq;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using System.Threading.Tasks;
using System.Threading;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Async;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that deals with feature activation and deactivation
    /// </summary>
    public static partial class FeatureExtensions
    {
        /// <summary>
        /// Activates a site collection or site scoped feature
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub web</param>
        /// <param name="featureID">ID of the feature to activate</param>
        /// <param name="sandboxed">Set to true if the feature is defined in a sandboxed solution</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static void ActivateFeature(this Web web, Guid featureID, bool sandboxed = false, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ActivateWebFeature, featureID);
            Task.Run(() => web.ProcessFeature(featureID, true, sandboxed, pollingIntervalSeconds)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Activates a site collection or site scoped feature
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub web</param>
        /// <param name="featureID">ID of the feature to activate</param>
        /// <param name="sandboxed">Set to true if the feature is defined in a sandboxed solution</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static async Task ActivateFeatureAsync(this Web web, Guid featureID, bool sandboxed = false, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ActivateWebFeature, featureID);
            await new SynchronizationContextRemover();
            await web.ProcessFeature(featureID, true, sandboxed, pollingIntervalSeconds);
        }

        /// <summary>
        /// Activates a site collection or site scoped feature
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="featureID">ID of the feature to activate</param>
        /// <param name="sandboxed">Set to true if the feature is defined in a sandboxed solution</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static void ActivateFeature(this Site site, Guid featureID, bool sandboxed = false, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ActivateWebFeature, featureID);
            Task.Run(() => site.ProcessFeature(featureID, true, sandboxed, pollingIntervalSeconds)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Activates a site collection or site scoped feature
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="featureID">ID of the feature to activate</param>
        /// <param name="sandboxed">Set to true if the feature is defined in a sandboxed solution</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static async Task ActivateFeatureAsync(this Site site, Guid featureID, bool sandboxed = false, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ActivateWebFeature, featureID);
            await new SynchronizationContextRemover();
            await site.ProcessFeature(featureID, true, sandboxed, pollingIntervalSeconds);
        }

        /// <summary>
        /// Deactivates a site collection or site scoped feature
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub web</param>
        /// <param name="featureID">ID of the feature to deactivate</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static void DeactivateFeature(this Web web, Guid featureID, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_DeactivateWebFeature, featureID);
            Task.Run(() => web.ProcessFeature(featureID, false, false, pollingIntervalSeconds)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Deactivates a site collection or site scoped feature
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub web</param>
        /// <param name="featureID">ID of the feature to deactivate</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static async Task DeactivateFeatureAsync(this Web web, Guid featureID, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_DeactivateWebFeature, featureID);
            await new SynchronizationContextRemover();
            await web.ProcessFeature(featureID, false, false, pollingIntervalSeconds);
        }

        /// <summary>
        /// Deactivates a site collection or site scoped feature
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="featureID">ID of the feature to deactivate</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static void DeactivateFeature(this Site site, Guid featureID, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_DeactivateWebFeature, featureID);
            Task.Run(() => site.ProcessFeature(featureID, false, false, pollingIntervalSeconds)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Deactivates a site collection or site scoped feature
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="featureID">ID of the feature to deactivate</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        public static async Task DeactivateFeatureAsync(this Site site, Guid featureID, int pollingIntervalSeconds = 30)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_DeactivateWebFeature, featureID);
            await new SynchronizationContextRemover();
            await site.ProcessFeature(featureID, false, false, pollingIntervalSeconds);
        }

        /// <summary>
        /// Checks if a feature is active
        /// </summary>
        /// <param name="site">Site to operate against</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <returns>True if active, false otherwise</returns>
        public static bool IsFeatureActive(this Site site, Guid featureID)
        {
            return Task.Run(() => IsFeatureActiveInternal(site.Features, featureID)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Checks if a feature is active
        /// </summary>
        /// <param name="site">Site to operate against</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <returns>True if active, false otherwise</returns>
        public static async  Task<bool> IsFeatureActiveAsync(this Site site, Guid featureID)
        {
            await new SynchronizationContextRemover();
            return await IsFeatureActiveInternal(site.Features, featureID);
        }

        /// <summary>
        /// Checks if a feature is active
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <returns>True if active, false otherwise</returns>
        public static bool IsFeatureActive(this Web web, Guid featureID)
        {
            return Task.Run(() => IsFeatureActiveInternal(web.Features, featureID)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Checks if a feature is active
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <returns>True if active, false otherwise</returns>
        public static async Task<bool> IsFeatureActiveAsync(this Web web, Guid featureID)
        {
            await new SynchronizationContextRemover();
            return await IsFeatureActiveInternal(web.Features, featureID);
        }

        /// <summary>
        /// Checks if a feature is active in the given FeatureCollection.
        /// </summary>
        /// <param name="features">FeatureCollection to check in</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <param name="noRetry">Use regular ExecuteQuery</param>
        /// <returns>True if active, false otherwise</returns>
        private static async Task<bool> IsFeatureActiveInternal(FeatureCollection features, Guid featureID, bool noRetry=false)
        {
            var featureIsActive = false;

            features.ClearObjectData();

            features.Context.Load(features);
            if (noRetry)
            {
                string clientTag = $"{PnPCoreUtilities.PnPCoreVersionTag}:ProcessFeatureInternal";
                if (clientTag.Length > 32)
                {
                    clientTag = clientTag.Substring(0, 32);
                }
                features.Context.ClientTag = clientTag;
                // Don't update this to ExecuteQueryRetry
                await features.Context.ExecuteQueryAsync();
            }
            else
            {
                await features.Context.ExecuteQueryRetryAsync();
            }

            var iprFeature = features.GetById(featureID);
            iprFeature.EnsureProperties(f => f.DefinitionId);

            if (iprFeature != null && iprFeature.IsPropertyAvailable("DefinitionId") && !iprFeature.ServerObjectIsNull.Value && iprFeature.DefinitionId.Equals(featureID))
            {
                featureIsActive = true;
            }

            return featureIsActive;
        }

        /// <summary>
        /// Activates or deactivates a site collection scoped feature
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="featureID">ID of the feature to activate/deactivate</param>
        /// <param name="activate">True to activate, false to deactivate the feature</param>
        /// <param name="sandboxed">Set to true if the feature is defined in a sandboxed solution</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        private static Task ProcessFeature(this Site site, Guid featureID, bool activate, bool sandboxed, int pollingIntervalSeconds = 30)
        {
            return ProcessFeatureInternal(site.Features, featureID, activate, sandboxed ? FeatureDefinitionScope.Site : FeatureDefinitionScope.Farm,pollingIntervalSeconds);
        }

        /// <summary>
        /// Activates or deactivates a web scoped feature
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub web</param>
        /// <param name="featureID">ID of the feature to activate/deactivate</param>
        /// <param name="activate">True to activate, false to deactivate the feature</param>
        /// <param name="sandboxed">True to specify that the feature is defined in a sandboxed solution</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        private static Task ProcessFeature(this Web web, Guid featureID, bool activate, bool sandboxed, int pollingIntervalSeconds = 30)
        {
            return ProcessFeatureInternal(web.Features, featureID, activate, sandboxed ? FeatureDefinitionScope.Site : FeatureDefinitionScope.Farm, pollingIntervalSeconds);
        }

        /// <summary>
        /// Activates or deactivates a site collection or web scoped feature
        /// </summary>
        /// <param name="features">Feature Collection which contains the feature</param>
        /// <param name="featureID">ID of the feature to activate/deactivate</param>
        /// <param name="activate">True to activate, false to deactivate the feature</param>
        /// <param name="scope">Scope of the feature definition</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        private static async Task ProcessFeatureInternal(FeatureCollection features, Guid featureID, bool activate, FeatureDefinitionScope scope, int pollingIntervalSeconds = 30)
        {
            if (activate)
            {
                // Feature enabling can take a long time, especially in case of the publishing feature...so let's make it more reliable
                features.Add(featureID, true, scope);

                if (pollingIntervalSeconds < 5)
                {
                    pollingIntervalSeconds = 5;
                }

                try
                {
                    string clientTag = $"{PnPCoreUtilities.PnPCoreVersionTag}:ProcessFeatureInternal";
                    if (clientTag.Length > 32)
                    {
                        clientTag = clientTag.Substring(0, 32);
                    }
                    features.Context.ClientTag = clientTag;
                    // Don't update this to ExecuteQueryRetry
                    await features.Context.ExecuteQueryAsync();
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ProcessFeatureInternal_FeatureActive, featureID);
                }
                catch (Exception ex)
                {
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ProcessFeatureInternal_FeatureException, ex.ToString());

                    // Don't wait for a "feature not found" exception, which is the typical exception we'll see
                    if (ex.HResult != -2146233088)
                    {
                        int retryAttempts = 10;
                        int retryCount = 0;

                        // wait and keep checking if the feature is active
                        while (retryAttempts > retryCount)
                        {
                            Thread.Sleep(TimeSpan.FromSeconds(pollingIntervalSeconds));
                            if (await IsFeatureActiveInternal(features, featureID, true))
                            {
                                retryCount = retryAttempts;
                                Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ProcessFeatureInternal_FeatureActivationState, true, featureID);
                            }
                            else
                            {
                                retryCount++;
                                Log.Info(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_ProcessFeatureInternal_FeatureActivationState, false, featureID);
                            }
                        }
                    }
                }
            }
            else
            {
                try
                {
                    features.Remove(featureID, false);
                    await features.Context.ExecuteQueryRetryAsync();
                }
                catch (Exception ex)
                {
                    Log.Error(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_FeatureActivationProblem, featureID, ex.Message);
                }
            }
        }
    }
}
