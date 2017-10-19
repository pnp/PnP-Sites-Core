using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class DependentFieldLookupIdToken : TokenDefinition
    {
        private Guid _actualId;

        public DependentFieldLookupIdToken(Web web, Guid idInTemplate, Guid actualId) :
            base(web, idInTemplate.ToString("D"))
        {
            this._actualId = actualId;
        }
        public override string GetReplaceValue()
        {
            return _actualId.ToString("D");
        }
    }
}
