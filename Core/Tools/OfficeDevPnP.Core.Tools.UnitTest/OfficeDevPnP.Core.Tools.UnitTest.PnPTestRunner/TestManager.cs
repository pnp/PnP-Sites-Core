using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.Linq;
using System.Management;
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

        /// <summary>
        /// Execute MSBuild run, collect the console output and store it in the database.
        /// </summary>
        /// <param name="pnpConfigurationToTest">Configuration to run</param>
        /// <param name="sqlConnectionString">Connection string to PnP Test Automation database that contains the configuration</param>
        /// <param name="buildTarget">Build file (.targets)</param>
        public void Execute(string pnpConfigurationToTest, string sqlConnectionString, string buildTarget)
        {
            // Prep an entity framework connection string and create a dbcontext
            connectionString = String.Format("metadata=res://*/SQL.TestModel.csdl|res://*/SQL.TestModel.ssdl|res://*/SQL.TestModel.msl;provider=System.Data.SqlClient;provider connection string=\"{0}\"", sqlConnectionString);
            context = new TestModelContainer(connectionString);

            string msBuildExe = ConfigurationManager.AppSettings[Constants.settingMSBuildExe].ToString();
            if (String.IsNullOrEmpty(msBuildExe))
            {
                msBuildExe = @"c:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe";
            }

            // Create process
            var proc = new Process();
            proc.StartInfo.FileName = msBuildExe;
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

            // wait for process to terminate gracefully...if not terminated after the defined timespan kill the process and all childprocesses spawned from this process
            int maxRunTimeInMinutes;
            if (!Int32.TryParse(ConfigurationManager.AppSettings[Constants.settingMaxRunTimeInMinutes], out maxRunTimeInMinutes))
            {
                maxRunTimeInMinutes = 180;
            }

            WriteLine(String.Format("Starting MSBuild with a max run time of {0} minutes", maxRunTimeInMinutes));
            if (!proc.WaitForExit(Convert.ToInt32(new TimeSpan(0, maxRunTimeInMinutes, 0).TotalMilliseconds)))
            {
                try
                {
                    WriteLine("[IMPORTANT] Started killing of the MSBuild process and it's child processes due to exceeded run time");

                    // kill vstest.console process if it's availabe
                    var vsTestProcesses = Process.GetProcessesByName("vstest.console");
                    if (vsTestProcesses.Length > 0)
                    {
                        for (int i = 0; i < vsTestProcesses.Length; i++)
                        {
                            KillAllProcessesSpawnedBy(Convert.ToUInt32(vsTestProcesses[i].Id));

                            WriteLine(string.Format("Killing {0}", vsTestProcesses[i].ProcessName));
                            vsTestProcesses[i].Kill();
                        }
                    }

                    // kill all the processes spawned by this process
                    KillAllProcessesSpawnedBy(Convert.ToUInt32(proc.Id));

                    // kill this process as it has been running too long
                    WriteLine(string.Format("Killing {0}", proc.ProcessName));
                    proc.Kill();
                    WriteLine("[IMPORTANT] Termination done!");
                }
                catch { }
            }

            // persist the log file
            TestRun run = context.TestRunSet.Find(this.testRunId);
            run.MSBuildLog = sb.ToString();
            // If the run did not finish by now then something went wrong
            if (run.Status != RunStatus.Done)
            {
                run.Status = RunStatus.Failed;
            }
            // Persist the changes
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
        /// Kill child processes of the provided process
        /// </summary>
        /// <param name="parentProcessId">Parent process ID</param>
        private void KillAllProcessesSpawnedBy(UInt32 parentProcessId)
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                    "SELECT * " +
                    "FROM Win32_Process " +
                    "WHERE ParentProcessId=" + parentProcessId);
                ManagementObjectCollection collection = searcher.Get();
                if (collection.Count > 0)
                {
                    foreach (var item in collection)
                    {
                        UInt32 childProcessId = (UInt32)item["ProcessId"];
                        if ((int)childProcessId != Process.GetCurrentProcess().Id)
                        {
                            Process childProcess = Process.GetProcessById((int)childProcessId);
                            WriteLine(string.Format("Killing {0}", childProcess.ProcessName));
                            childProcess.Kill();
                        }
                    }
                }
            }
            catch
            { }
        }

        /// <summary>
        /// Write data to screen and stringbuilder
        /// </summary>
        /// <param name="s">String to write</param>
        private void WriteLine(string s)
        {
            sb.AppendLine(s);
            Console.WriteLine(s);
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
