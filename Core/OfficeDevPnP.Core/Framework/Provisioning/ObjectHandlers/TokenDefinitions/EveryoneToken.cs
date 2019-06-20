#if !ONPREMISES
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{everyone}",
       Description = "Returns the claim for everyone in this tenant",
       Example = "{everyone}",
       Returns = "c:0(.s|true")]
    internal class EveryoneToken : TokenDefinition
    {

        public EveryoneToken(Web web)
            : base(web, $"{{everyone}}")
        {

        }

        public override string GetReplaceValue()
        {
            return "c:0(.s|true";
        }
    }
}
#endif