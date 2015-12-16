using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Operations
{
    /// <summary>
    /// Register an app using appregnew.aspx
    /// </summary>
    public class AppRegNew : RemoteOperation
    {
        #region Construction
        public AppRegNew(string TargetUrl, AuthenticationType authType, string User, SecureString Password, string AppInstance = "", string Domain = "")
            : base(TargetUrl, authType, User, Password, AppInstance, Domain)
        {
        }
        #endregion

        #region Properties
        public override string OperationPageUrl
        {
            get
            {
                return "/_layouts/15/appregnew.aspx";
            }
        }

        public string AppId
        {
            get;
            set;
        }

        public string AppSecret
        {
            get;
            set;
        }

        public string Title
        {
            get;
            set;
        }

        public string HostUri
        {
            get;
            set;
        }

        public string RedirectUri
        {
            get;
            set;
        }
        #endregion

        #region Methods
        public override void SetPostVariablesOnline()
        {
            // Set operation specific parameters for online
            this.PostParameters.Add("__EVENTTARGET", "ctl00$PlaceHolderMain$ctl01$RptControls$BtnCreate");

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$AppTypeControl$RadAppType", "RadAppTypeExternal");

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl01$TxtAppId", AppId);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$AppSecretControl$TxtAppSecret", AppSecret);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl02$TxtTitle", Title);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl03$TxtHostUri", HostUri);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl04$TxtRedirectUri", RedirectUri);
        }

        public override void SetPostVariablesOnPremise()
        {
            // Set operation specific parameters for on-premise
            this.PostParameters.Add("__EVENTTARGET", "ctl00$PlaceHolderMain$ctl01$RptControls$BtnCreate");

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$AppTypeControl$RadAppType", "RadAppTypeExternal");

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl01$TxtAppId", AppId);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl02$TxtAppSecret", AppSecret);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl03$TxtTitle", Title);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl04$TxtHostUri", HostUri);

            this.PostParameters.Add("ctl00$PlaceHolderMain$AppInfoSection$ctl05$TxtRedirectUri", RedirectUri);
        }

        public static string GenerateAppSecret()
        {
            Random rnd = new Random();
            byte[] AppSecret = new byte[32];
            for (int i = 0; i < 32; i++)
                AppSecret[i] = (byte)rnd.Next(256);
            return System.Convert.ToBase64String(AppSecret);
        }

        public static string GenerateAppId()
        {
            return Guid.NewGuid().ToString("D"); 
        }

        #endregion
    }
}
