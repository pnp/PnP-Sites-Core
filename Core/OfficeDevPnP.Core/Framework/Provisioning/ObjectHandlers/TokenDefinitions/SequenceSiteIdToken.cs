#if !ONPREMISES
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
    Token = "{sequencesiteid:[provisioningid]}",
    Description = "Returns a id of the site given its provisioning ID from the sequence",
    Example = "{sequencesiteid:MYID}",
    Returns = "https://contoso.sharepoint.com/sites/mynewsite")]
    internal class SequenceSiteIdToken : TokenDefinition
    {
        private Guid _id = Guid.Empty;
        public SequenceSiteIdToken(Web web, string provisioningId, Guid id)
            : base(web, $"{{sequencesiteid:{provisioningId}}}")
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