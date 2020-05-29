#if !ONPREMISES
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
    Token = "{sequencesitegroupid:[provisioningid]}",
    Description = "Returns a Id of the associated group or an empty guid if not O365 has been associated with this site",
    Example = "{sequencesitegroupid:MYID}",
    Returns = "c7d9f9aa-4696-4c27-8a22-7d8eb7e70fda")]
    internal class SequenceSiteGroupIdToken : TokenDefinition
    {
        private Guid _id = Guid.Empty;
        public SequenceSiteGroupIdToken(Web web, string provisioningId, Guid id)
            : base(web, $"{{sequencesitegroupid:{provisioningId}}}")
        {
            _id = id;
        }

        public override string GetReplaceValue()
        {
            return _id.ToString();
        }
    }
}
#endif