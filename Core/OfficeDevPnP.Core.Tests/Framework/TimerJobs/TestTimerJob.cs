using OfficeDevPnP.Core.Framework.TimerJobs;
using System;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.TimerJobs
{
    class TestTimerJob : TimerJob
    {
        public TestTimerJob(string name) : base(name)
        {
            TimerJobRun += TestTimerJob_TimerJobRun;
        }

        private void TestTimerJob_TimerJobRun(object sender, TimerJobRunEventArgs e)
        {
            var web = e.SiteClientContext.Web;

            var users = e.SiteClientContext.LoadQuery(web.SiteUsers.Include(uc => uc.Email).Where(uc => uc.IsSiteAdmin));

            e.SiteClientContext.ExecuteQueryRetry();
        }
    }
}
