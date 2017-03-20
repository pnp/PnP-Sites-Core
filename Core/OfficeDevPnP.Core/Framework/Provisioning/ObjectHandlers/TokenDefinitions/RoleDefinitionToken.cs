using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class RoleDefinitionToken : TokenDefinition
    {
        private string name;
        public RoleDefinitionToken(Web web, RoleDefinition definition)
            : base(web, $"{{roledefinition:{definition.RoleTypeKind}}}")
        {
            name = definition.EnsureProperty(r => r.Name);
            
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = name;
            }
            return CacheValue;
        }
    }
}