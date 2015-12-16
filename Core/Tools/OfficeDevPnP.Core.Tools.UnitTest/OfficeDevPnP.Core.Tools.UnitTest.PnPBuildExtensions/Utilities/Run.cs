using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Utilities
{
    public class Run
    {
        /// <summary>
        /// Method to invoke ps1 files from c#
        /// </summary>
        /// <param name="powershellpath">path of ps1 to be execute</param>
        /// <param name="Parameters">collection of paramets/arguments in the ps1 script file</param>
        /// <returns></returns>
        public static string RunScript(string powershellpath, Hashtable Parameters)
        {
            StringBuilder stringBuilder = new StringBuilder();
            try
            {
                // create Powershell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                //RunspaceInvoke runSpaceInvoker = new RunspaceInvoke(runspace);
                //runSpaceInvoker.Invoke("Set-ExecutionPolicy Unrestricted");

                // create a pipeline and feed it the script text
                Pipeline pipeline = runspace.CreatePipeline();
                Command command = new Command(powershellpath);

                //Looping the parametrs
                foreach (DictionaryEntry Parameter in Parameters)
                {
                    command.Parameters.Add(Convert.ToString(Parameter.Key), Convert.ToString(Parameter.Value));
                }
                pipeline.Commands.Add(command);

                Collection<PSObject> results = pipeline.Invoke();
                runspace.Close();

                // convert the script result into a single string 
                foreach (PSObject obj in results)
                {
                    stringBuilder.AppendLine(obj.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }

            return stringBuilder.ToString();
        }
    }
}
