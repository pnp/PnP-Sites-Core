using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    class FieldValidator: ValidatorBase
    {
        public bool Validate(FieldCollection sourceCollection, FieldCollection targetCollection, TokenParser tokenParser)
        {
            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            parserSettings.Add("SchemaXml", new string[] { "~sitecollection", "~site", "{sitecollectiontermstoreid}", "{termsetid}" });
            bool isFieldMatch = ValidateObjectsXML(sourceCollection, targetCollection, "SchemaXml", new List<string> { "ID" }, tokenParser, parserSettings);
            Console.WriteLine("-- Field validation " + isFieldMatch);
            return isFieldMatch;
        }
    }
}
