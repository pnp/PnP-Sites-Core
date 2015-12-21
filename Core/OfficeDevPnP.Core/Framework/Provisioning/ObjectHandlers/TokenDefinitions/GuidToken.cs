using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class GuidToken : TokenDefinition
    {
        public GuidToken(Web web)
            : base(web, "~guid", "{guid}")
        {
        }

        public override string GetReplaceValue()
        {
            return Guid.NewGuid().ToString();
        }
    }
}