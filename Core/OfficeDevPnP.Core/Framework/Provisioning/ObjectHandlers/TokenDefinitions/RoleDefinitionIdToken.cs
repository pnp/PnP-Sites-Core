using Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
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
