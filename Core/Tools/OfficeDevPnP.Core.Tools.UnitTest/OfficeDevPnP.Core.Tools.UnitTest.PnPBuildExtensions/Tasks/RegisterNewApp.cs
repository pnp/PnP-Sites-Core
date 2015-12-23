using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;


namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Tasks
{
    public class RegisterNewApp : Task
    {
        public String CredentialManagerLabel
        {
            get;
            set;
        }

        public String UserName
        {
            get;
            set;
        }

        public String Password
        {
            get;
            set;
        }

        [Required]
        [Output]
        public String SharePointSiteUrl
        {
            get;
            set;
        }

        [Output]
        public String ClientId
        {
            get;
            set;
        }

        [Output]
        public String ClientSecret
        {
            get;
            set;
        }

        [Required]
        [Output]
        public String Title
        {
            get;
            set;
        }

        [Required]
        [Output]
        public String AppDomain
        {
            get;
            set;
        }

        [Required]
        [Output]
        public String RedirectUri
        {
            get;
            set;
        }

        public override bool Execute()
        {
            try
            {
                Log.LogMessageFromText(String.Format("RegisterNewApp: Registering new app with title {0} for site {1}", Title, SharePointSiteUrl), MessageImportance.Normal);

                //DebugBreak();

                AppManager am;
                if (!string.IsNullOrEmpty(CredentialManagerLabel))
                {
                    am = new AppManager(SharePointSiteUrl, CredentialManagerLabel);
                }
                else
                {
                    var spoPassword = new SecureString();
                    foreach (char c in this.Password)
                    {
                        spoPassword.AppendChar(c);
                    }
                    am = new AppManager(SharePointSiteUrl, UserName, spoPassword);
                }

                string clientId = ClientId;
                string clientSecret = ClientSecret;
                if (am.RegisterApplication(ref clientId, ref clientSecret, Title, AppDomain, RedirectUri))
                {
                    ClientId = clientId;
                    ClientSecret = clientSecret;
                    return true;
                }
                else
                {
                    Log.LogError("RegisterNewApp: Registering new app with title {0} for site {1} failed", Title, SharePointSiteUrl);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
                return false;
            }
        }

        #region Private methods
        [Conditional("DEBUG")]
        void DebugBreak()
        {
            Debugger.Launch();
        }
        #endregion

    }
}
