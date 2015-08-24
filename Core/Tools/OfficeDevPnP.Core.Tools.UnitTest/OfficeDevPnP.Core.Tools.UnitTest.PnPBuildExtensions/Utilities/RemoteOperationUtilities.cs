using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Utilities
{
    class RemoteOperationUtilities
    {
        public static string ReadHiddenField(string pageHtml, string fieldName)
        {
            string result = "";
            string hiddenFieldFlag = string.Format("id=\"{0}\" value=\"", fieldName);
            int i = pageHtml.IndexOf(hiddenFieldFlag);
            if (i > -1)
            {
                i = i + hiddenFieldFlag.Length;
                int j = pageHtml.IndexOf("\"", i);
                result = HttpUtility.UrlEncode(pageHtml.Substring(i, j - i));
            }

            return result;
        }

        public static string ReadInputFieldById(string pageHtml, string fieldName)
        {
            string result = "";
            string inputFieldFlag = string.Format("id=\"{0}\"", fieldName);
            int i = pageHtml.IndexOf(inputFieldFlag);
            if (i > -1)
            {
                i = i + inputFieldFlag.Length;
                int j = pageHtml.IndexOf("value=\"", i) + "value=\"".Length;
                int k = pageHtml.IndexOf("\"", j);
                result = HttpUtility.UrlEncode(pageHtml.Substring(j, k - j));
            }

            return result;
        }

        public static string FormatOperationUrlString(string hostUrl, string operationPageUrl)
        {
            return hostUrl.TrimEnd(new char[] { '/' }) + operationPageUrl;
        }
    }
}
