using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    public class ValidateXmlEventArgs: EventArgs
    {
        public XElement SourceObject { get; set; }
        public XElement TargetObject { get; set; }
        public bool IsEqual { get; set; }


        public ValidateXmlEventArgs(XElement sourceObject, XElement targetObject)
        {
            SourceObject = sourceObject;
            TargetObject = targetObject;
            IsEqual = false;
        }

    }
}
