using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{

    [TokenDefinitionDescription(
      Token = "{groupsitename}",
      Description = "Returns the title of the current site with special char \" / \\ [ ] : | < > + = ; , ? * ' @ replaced with _",
      Example = "{groupsitename}",
      Returns = "Titel with replaced special char")]
    internal class GroupSiteTitleToken : VolatileTokenDefinition
    {
        public GroupSiteTitleToken(Web web) : base(web, "{groupsitetitle}", "{groupsitename}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Load(TokenContext.Web, w => w.Title);
                TokenContext.ExecuteQueryRetry();
                CacheValue = Regex.Replace(TokenContext.Web.Title, "[\"/\\[\\]\\\\:|<>+=;,?*\'@]", "_"); 
            }
            return CacheValue;
        }
    }

}
