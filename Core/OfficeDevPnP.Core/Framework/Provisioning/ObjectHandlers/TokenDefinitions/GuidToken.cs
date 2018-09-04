using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{guid}",
     Description = "Returns a newly generated GUID",
     Example = "{guid}",
     Returns = "f2cd6d5b-1391-480e-a3dc-7f7f96137382")]
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