using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{fieldid:[internalname]}",
     Description = "Returns the ID of a field given its internalname",
     Example = "{fieldid:LeaveEarly}",
     Returns = "20d5ad60-8662-4d06-92bb-3a434766f344")]
    internal class FieldIdToken : TokenDefinition
    {
        private readonly string _value = null;

        public FieldIdToken(Web web, string InternalName, System.Guid fieldId)
            : base(web, $"{{fieldid:{InternalName}}}")
        {
            _value = fieldId.ToString();
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