using System.Collections.Generic;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildTest
{
    class Stats : ITestRunStatistics
    {
        public long ExecutedTests
        {
            get { return 234; }
        }

        IDictionary<TestOutcome, long> ITestRunStatistics.Stats => throw new System.NotImplementedException();

        public long this[Microsoft.VisualStudio.TestPlatform.ObjectModel.TestOutcome testOutcome]
        {
            get { return 0; }
        }
    }
}
