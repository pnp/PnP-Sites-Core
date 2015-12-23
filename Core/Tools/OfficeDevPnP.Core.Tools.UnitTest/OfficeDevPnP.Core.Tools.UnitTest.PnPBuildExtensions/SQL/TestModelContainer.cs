using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL
{

    /// <summary>
    /// Partial class to offer a connectionstring override
    /// </summary>
    public partial class TestModelContainer : DbContext
    {
        public TestModelContainer(string connectionString)
            : base(connectionString)
        {
        }
    }
}
