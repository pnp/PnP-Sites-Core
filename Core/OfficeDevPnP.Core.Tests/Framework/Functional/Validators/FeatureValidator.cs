using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    public static class FeatureValidator
    {
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

        public static bool ValidateFeatures(FeatureCollection sFeatures, FeatureCollection tFeatures)
        {
            int sCount = 0;
            int tCount = 0;

            foreach (Feature sFeature in sFeatures)
            {
                sCount++;
                Guid sID = sFeature.Id;

                Feature tFeature = tFeatures.Where(ft => ft.Id == sID).FirstOrDefault();
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
