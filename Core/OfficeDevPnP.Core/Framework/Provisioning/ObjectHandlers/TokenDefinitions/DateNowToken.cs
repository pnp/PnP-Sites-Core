using System;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{now}",
       Description = "Returns the current date in universal date time format: yyyy-MM-ddTHH:mm:ss.fffK",
       Example = "{now}",
       Returns = "2018-04-18T15:44:45.898+02:00")]
    /// <summary>
    /// Gets current date time in universal date time format yyyy-MM-ddTHH:mm:ss.fffK 
    /// </summary>
    public class DateNowToken : TokenDefinition
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="web">A SharePoint site/subsite</param>
        public DateNowToken(Web web)
            : base(web, "~now", "{now}")
        {
        }

        /// <summary>
        /// replaces current date time to the universal date time format yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffK
        /// </summary>
        /// <returns>date time in yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffK format</returns>
        public override string GetReplaceValue()
        {
            return DateTime.Now.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffK");
        }
    }
}
