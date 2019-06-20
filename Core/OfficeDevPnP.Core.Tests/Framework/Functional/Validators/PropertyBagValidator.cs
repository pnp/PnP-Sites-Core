using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    public class PropertyBagValidator: ValidatorBase
    {
        public bool Validate(PropertyBagEntryCollection sourceCollection, PropertyBagEntryCollection targetCollection, TokenParser tokenParser)
        {
            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            parserSettings.Add("Value", null);
            bool isPropertyBagsMatch = ValidateObjects(sourceCollection, targetCollection, new List<string> { "Key", "Value", "Indexed" }, tokenParser, parserSettings);
            Console.WriteLine("-- Property Bags validation " + isPropertyBagsMatch);
            return isPropertyBagsMatch; 
        }
    }
}
