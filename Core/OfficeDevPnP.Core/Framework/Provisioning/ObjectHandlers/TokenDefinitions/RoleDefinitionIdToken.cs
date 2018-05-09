using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{roledefinitionid:[rolename]}",
        Description = "Returns the id of the given role definition name",
        Example = "{roledefinitionid:My Role Definition}",
        Returns = "23")]
    internal class RoleDefinitionIdToken : TokenDefinition
    {
        private readonly int _roleDefinitionId = 0;
        public RoleDefinitionIdToken(Web web, string name, int roleDefinitionId)
            : base(web, $"{{roledefinitionid:{Regex.Escape(name)}}}")
        {
            _roleDefinitionId = roleDefinitionId;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _roleDefinitionId.ToString();
            }
            return CacheValue;
        }
    }
}
