using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length <= 2)
            {
                Console.WriteLine("Usage: OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner.exe testconfiguration sqlconnectionstring pathtotargetfile");
                return;
            }

            TestManager tm = new TestManager();
            tm.Execute(args[0], args[1], args[2]);
            //tm.Execute("BertMTFirstReleaseCredentials", 
            //           "data source=(localdb)\\MSSQLLocalDB;initial catalog=PnPTestAutomation;integrated security=True;MultipleActiveResultSets=True;App=EntityFramework", 
            //           "C:\\GitHub\\BertPnPSitesCore\\Core\\Tools\\OfficeDevPnP.Core.Tools.UnitTest\\PnPSQLCore.targets");

            Console.ReadLine();
        }


    }
}
