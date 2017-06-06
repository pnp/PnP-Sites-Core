using System;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    public class DateNowToken : TokenDefinition
    {
        public DateNowToken(Web web)
            : base(web, "~now", "{now}")
        {
        }

        public override string GetReplaceValue()
        {
            return DateTime.Now.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffK");
        }
    }
}
