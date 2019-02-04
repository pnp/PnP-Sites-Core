#if !ONPREMISES
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
    Token = "{sequencesitecollectionid:[provisioningid]}",
    Description = "Returns a site collection id of the site given its provisioning ID from the sequence",
    Example = "{sequencesitecollectionid:MYID}",
    Returns = "https://contoso.sharepoint.com/sites/mynewsite")]
    internal class SequenceSiteCollectionIdToken : TokenDefinition
    {
        private Guid _id = Guid.Empty;
        public SequenceSiteCollectionIdToken(Web web, string provisioningId, Guid id)
            : base(web, $"{{sequencesitecollectionid:{provisioningId}}}")
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