using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    /// <summary>
    /// Feature validation class
    /// </summary>
    public class FeatureValidator : ValidatorBase
    {
        /// <summary>
        /// Validate site and web scoped features
        /// </summary>
        /// <param name="source">Source features</param>
        /// <param name="target">Target features</param>
        /// <returns>True if both source and target features match, false otherwise</returns>
        public static bool Validate(Features source, Features target)
        {
            bool isSiteFeaturesMatch = ValidateFeatures(source.SiteFeatures, target.SiteFeatures);
            Console.WriteLine("Site Features validation " + isSiteFeaturesMatch);

            bool isWebFeaturesMatch = ValidateFeatures(source.WebFeatures, target.WebFeatures);
            Console.WriteLine("Web Features validation " + isWebFeaturesMatch);

            if (!isSiteFeaturesMatch || !isWebFeaturesMatch)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Compare source and target features
        /// </summary>
        /// <param name="sFeatures">Source features</param>
        /// <param name="tFeatures">Target features</param>
        /// <returns>True if the features match, false otherwise</returns>
        public static bool ValidateFeatures(FeatureCollection sFeatures, FeatureCollection tFeatures)
        {
            int sCount = 0;
            int tCount = 0;

            foreach (Feature sFeature in sFeatures)
            {
                sCount++;
                Guid sID = sFeature.Id;

                Feature tFeature = tFeatures.Where(ft => ft.Id == sID).FirstOrDefault();
                
                // Feature activation: do we see the target feature with the correct id?
                // Feature deactivation: we shouldn't see the target feature anymore when we choose to deactivate
                if ((tFeature != null && sID == tFeature.Id) || (sFeature.Deactivate && tFeature == null))
                {
                    tCount++;
                }
            }

            if (sCount != tCount)
            {
                return false;
            }

            return true;
        }

    }
}
