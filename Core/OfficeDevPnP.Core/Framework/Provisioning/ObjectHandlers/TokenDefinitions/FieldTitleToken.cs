using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class FieldTitleToken : TokenDefinition
    {
        private readonly string _value = null;
        public FieldTitleToken(Web web, string InternalName, string Title)
            : base(web, $"{{fieldtitle:{InternalName}}}")
        {
            _value = Title;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}