using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class CDATAStartToken : TokenDefinition
    {
        public CDATAStartToken(Web web)
            : base(web, "~cdatastart", "{cdatastart}")
        {
        }

        public override string GetReplaceValue()
        {
            return "<![CDATA[";
        }
    }
}