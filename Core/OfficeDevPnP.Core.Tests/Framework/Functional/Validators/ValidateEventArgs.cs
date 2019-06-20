using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    public class ValidateEventArgs: EventArgs
    {
        public string PropertyName { get; set; }
        public string SourceValue { get; set; }
        public string TargetValue { get; set; }
        public object SourceObject { get; set; }
        public object TargetObject { get; set; }
        public bool IsEqual { get; set; }


        public ValidateEventArgs(string propertyName, string sourceValue, string targetValue, object sourceObject, object targetObject)
        {
            PropertyName = propertyName;
            SourceValue = sourceValue;
            TargetValue = targetValue;
            SourceObject = sourceObject;
            TargetObject = targetObject;
            IsEqual = false;
        }

    }
}
