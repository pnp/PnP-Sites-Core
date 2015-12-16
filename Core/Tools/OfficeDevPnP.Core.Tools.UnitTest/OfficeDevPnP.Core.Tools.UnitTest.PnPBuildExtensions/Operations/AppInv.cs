using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Operations
{
    /// <summary>
    /// Trust the app appinv.aspx
    /// </summary>
    public class AppInv : RemoteOperation
    {
        #region Construction
        public AppInv(string TargetUrl, AuthenticationType authType, string User, SecureString Password, string AppInstanceId, string Domain = "")
            : base(TargetUrl, authType, User, Password, AppInstanceId, Domain)
        {
        }
        #endregion

        #region Properties
        public override string OperationPageUrl
        {
            get
            {
                return "/_layouts/15/appinv.aspx?AppInstanceId=" + AppInstanceId;
            }
        }
        #endregion

        #region Methods
        public override void SetPostVariablesOnline()
        {
            // Set operation specific parameters for online
            this.PostParameters.Add("__EVENTTARGET", "ctl00$PlaceHolderMain$BtnAllow");
        }

        public override void SetPostVariablesOnPremise()
        {
            // Set operation specific parameters for on-premise
            this.PostParameters.Add("__EVENTTARGET", "ctl00$PlaceHolderMain$BtnAllow");

        }

        public static string GetRandomAppSecret()
        {
            Random rnd = new Random();
            byte[] AppSecret = new byte[32];
            for (int i = 0; i < 32; i++)
                AppSecret[i] = (byte)rnd.Next(256);
            return System.Convert.ToBase64String(AppSecret);
        }
        #endregion
    }
}

