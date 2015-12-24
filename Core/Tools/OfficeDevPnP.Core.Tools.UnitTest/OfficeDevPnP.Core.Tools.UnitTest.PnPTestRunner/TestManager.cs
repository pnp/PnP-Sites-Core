using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPTestRunner
{
    public class TestManager
    {
        private StringBuilder sb;
        private string connectionString;
        private TestModelContainer context;
        private int testRunId;

        public TestManager()
        {
            this.sb = new StringBuilder(500);
        }

        public void Execute(string pnpConfigurationToTest, string sqlConnectionString, string buildTarget)
        {
            // Prep an entity framework connection string and create a dbcontext
            connectionString = String.Format("metadata=res://*/SQL.TestModel.csdl|res://*/SQL.TestModel.ssdl|res://*/SQL.TestModel.msl;provider=System.Data.SqlClient;provider connection string=\"{0}\"", sqlConnectionString);
            context = new TestModelContainer(connectionString);

            // Create process
            var proc = new Process();
            proc.StartInfo.FileName = @"c:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe";
            proc.StartInfo.Arguments = String.Format("/property:PnPConfigurationToTest={0};PnPSQLConnectionString=\"metadata=res://*/SQL.TestModel.csdl|res://*/SQL.TestModel.ssdl|res://*/SQL.TestModel.msl;provider=System.Data.SqlClient;provider connection string=&quot;{1}&quot;\"  \"{2}\"", pnpConfigurationToTest, sqlConnectionString, buildTarget);

            // set up output redirection
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;
            proc.EnableRaisingEvents = true;
            proc.StartInfo.CreateNoWindow = true;
            proc.StartInfo.UseShellExecute = false;
            // see below for output handler
            proc.ErrorDataReceived += proc_DataReceived;
            proc.OutputDataReceived += proc_DataReceived;

            proc.Start();

            proc.BeginErrorReadLine();
            proc.BeginOutputReadLine();

            // wait for process to terminate
            proc.WaitForExit();

            // persist the log file
            TestRun run = context.TestRunSet.Find(this.testRunId);
            run.MSBuildLog = sb.ToString();
            SaveChanges();
        }

        void proc_DataReceived(object sender, DataReceivedEventArgs e)
        {
            // output will be in string e.Data
            if (e.Data != null)
            {
                sb.AppendLine(e.Data);
                
                // anyhow output the results in case of an interactive run
                Console.WriteLine(e.Data);

                if (e.Data.Contains("[PnPTestRunID:"))
                {
                    this.testRunId = Convert.ToInt32(e.Data.Replace("[PnPTestRunID:", "").Replace("]", ""));
                }
            }
        }

        /// <summary>
        /// Persists changes using the entity framework. Puts detailed DbEntityValidationException errors in console
        /// </summary>
        private void SaveChanges()
        {
            try
            {
                context.SaveChanges();
            }
            catch (DbEntityValidationException e)
            {
                foreach (var eve in e.EntityValidationErrors)
                {
                    Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                        eve.Entry.Entity.GetType().Name, eve.Entry.State);
                    foreach (var ve in eve.ValidationErrors)
                    {
                        Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"",
                            ve.PropertyName, ve.ErrorMessage);
                    }
                }
                throw;
            }
        }
    }
}
