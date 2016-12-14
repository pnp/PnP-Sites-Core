using System;
using System.Linq;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using System.Threading.Tasks;
using System.Threading;
using OfficeDevPnP.Core.Utilities;

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
            web.ProcessFeature(featureID, true, sandboxed, pollingIntervalSeconds);
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
            site.ProcessFeature(featureID, true, sandboxed, pollingIntervalSeconds);
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
            web.ProcessFeature(featureID, false, false, pollingIntervalSeconds);
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
            site.ProcessFeature(featureID, false, false, pollingIntervalSeconds);
        }

        /// <summary>
        /// Checks if a feature is active
        /// </summary>
        /// <param name="site">Site to operate against</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <returns>True if active, false otherwise</returns>
        public static bool IsFeatureActive(this Site site, Guid featureID)
        {
            return IsFeatureActiveInternal(site.Features, featureID);
        }

        /// <summary>
        /// Checks if a feature is active
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <returns>True if active, false otherwise</returns>
        public static bool IsFeatureActive(this Web web, Guid featureID)
        {
            return IsFeatureActiveInternal(web.Features, featureID);
        }

        /// <summary>
        /// Checks if a feature is active in the given FeatureCollection.
        /// </summary>
        /// <param name="features">FeatureCollection to check in</param>
        /// <param name="featureID">ID of the feature to check</param>
        /// <param name="noRetry">Use regular ExecuteQuery</param>
        /// <returns>True if active, false otherwise</returns>
        private static bool IsFeatureActiveInternal(FeatureCollection features, Guid featureID, bool noRetry=false)
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
                features.Context.ExecuteQuery();
            }
            else
            {
                features.Context.ExecuteQueryRetry();
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
        private static void ProcessFeature(this Site site, Guid featureID, bool activate, bool sandboxed, int pollingIntervalSeconds = 30)
        {
            ProcessFeatureInternal(site.Features, featureID, activate, sandboxed ? FeatureDefinitionScope.Site : FeatureDefinitionScope.Farm,pollingIntervalSeconds);
        }

        /// <summary>
        /// Activates or deactivates a web scoped feature
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub web</param>
        /// <param name="featureID">ID of the feature to activate/deactivate</param>
        /// <param name="activate">True to activate, false to deactivate the feature</param>
        /// <param name="sandboxed">True to specify that the feature is defined in a sandboxed solution</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        private static void ProcessFeature(this Web web, Guid featureID, bool activate, bool sandboxed, int pollingIntervalSeconds = 30)
        {
            ProcessFeatureInternal(web.Features, featureID, activate, sandboxed ? FeatureDefinitionScope.Site : FeatureDefinitionScope.Farm, pollingIntervalSeconds);
        }

        /// <summary>
        /// Activates or deactivates a site collection or web scoped feature
        /// </summary>
        /// <param name="features">Feature Collection which contains the feature</param>
        /// <param name="featureID">ID of the feature to activate/deactivate</param>
        /// <param name="activate">True to activate, false to deactivate the feature</param>
        /// <param name="scope">Scope of the feature definition</param>
        /// <param name="pollingIntervalSeconds">The time in seconds between polls for "IsActive"</param>
        private static void ProcessFeatureInternal(FeatureCollection features, Guid featureID, bool activate, FeatureDefinitionScope scope, int pollingIntervalSeconds = 30)
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
                    features.Context.ExecuteQuery();
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
                            if (IsFeatureActiveInternal(features, featureID, true))
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
                    features.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    Log.Error(Constants.LOGGING_SOURCE, CoreResources.FeatureExtensions_FeatureActivationProblem, featureID, ex.Message);
                }
            }
        }
    }
}
