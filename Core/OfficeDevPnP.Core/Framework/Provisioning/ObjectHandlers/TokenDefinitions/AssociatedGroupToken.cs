using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{associatedownersgroup}",
       Description = "Returns the title of the associated owners SharePoint group of a site",
       Example = "{associatedownersgroup}",
       Returns = "My Site Owners Group Title")]
    [TokenDefinitionDescription(
       Token = "{associatedmembersgroup}",
       Description = "Returns the title of the associated members SharePoint group of a site",
       Example = "{associatedmembersgroup}",
       Returns = "My Site Members Group Title")]
    [TokenDefinitionDescription(
       Token = "{associatedvisitorsgroup}",
       Description = "Returns the title of the associated visitors SharePoint group of a site",
       Example = "{associatedvisitorsgroup}",
       Returns = "My Site Visitors Group Title")]
    internal class AssociatedGroupToken : TokenDefinition
    {
        private AssociatedGroupType _groupType;

        public AssociatedGroupToken(Web web, AssociatedGroupType groupType)
            : base(web, $"{{associated{groupType.ToString().TrimEnd('s')}group}}")
        {
            _groupType = groupType;
        }

        public override string GetReplaceValue()
        {

            if (string.IsNullOrEmpty(CacheValue))
            {
                TokenContext.Load(TokenContext.Web, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                TokenContext.ExecuteQueryRetry();
                switch (_groupType)
                {
                    case AssociatedGroupType.owners:
                        {
                            CacheValue = TokenContext.Web.AssociatedOwnerGroup.Title;
                            break;
                        }
                    case AssociatedGroupType.members:
                        {
                            CacheValue = TokenContext.Web.AssociatedMemberGroup.Title;
                            break;
                        }
                    case AssociatedGroupType.visitors:
                        {
                            CacheValue = TokenContext.Web.AssociatedVisitorGroup.Title;
                            break;
                        }
                }
            }
            return CacheValue;
        }

        public enum AssociatedGroupType
        {
            owners,
            members,
            visitors
        }
    }
}