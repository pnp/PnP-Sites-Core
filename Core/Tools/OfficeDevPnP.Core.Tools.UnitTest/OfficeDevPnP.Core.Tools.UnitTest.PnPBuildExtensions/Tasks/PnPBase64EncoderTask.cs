using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Tasks
{
    public class PnPBase64EncoderTask : Task
    {
        [Required]
        public String Input
        {
            get;
            set;
        }

        [Output]
        public string Output
        {
            get;
            set;
        }

        public override bool Execute()
        {
            try
            {
                //Log.LogMessageFromText(String.Format("Base64 encode string {0}", Input), MessageImportance.Normal);
                Output = Base64Encode(Input).Replace("=", "&equal");
                return true;
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
                return false;
            }
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }
    }
}
