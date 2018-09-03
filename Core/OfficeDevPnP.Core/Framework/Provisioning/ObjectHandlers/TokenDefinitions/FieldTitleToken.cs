using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{fieldtitle:[internalname]}",
     Description = "Returns the title/displayname of a field given its internalname",
     Example = "{fieldtitle:LeaveEarly}",
     Returns = "Leaving Early")]
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