using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class CDATAEndToken : TokenDefinition
    {
        public CDATAEndToken(Web web)
            : base(web, "~cdataend", "{cdataend}")
        {
        }

        public override string GetReplaceValue()
        {
            return "]]>";
        }
    }
}