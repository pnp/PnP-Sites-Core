using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{associatedownergroupid}",
       Description = "Returns the id of the associated owners SharePoint group of a site",
       Example = "{associatedownergroupid}",
       Returns = "My Site Owners Group id")]
    [TokenDefinitionDescription(
       Token = "{associatedmembergroupid}",
       Description = "Returns the id of the associated members SharePoint group of a site",
       Example = "{associatedmembergroupid}",
       Returns = "My Site Members Group id")]
    [TokenDefinitionDescription(
       Token = "{associatedvisitorgroupid}",
       Description = "Returns the id of the associated visitors SharePoint group of a site",
       Example = "{associatedvisitorgroupid}",
       Returns = "My Site Visitors Group id")]
    internal class AssociatedGroupIdToken : VolatileTokenDefinition
    {
        private AssociatedGroupType _groupType;

        public AssociatedGroupIdToken(Web web, AssociatedGroupType groupType)
            : base(web, $"{{associated{groupType.ToString().TrimEnd('s')}groupid}}")
        {
            _groupType = groupType;
        }

        public override string GetReplaceValue()
        {

            if (string.IsNullOrEmpty(CacheValue))
            {
                switch (_groupType)
                {
                    case AssociatedGroupType.owners:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedOwnerGroup.Id);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedOwnerGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedOwnerGroup.Id.ToString();
                            }
                            break;
                        }
                    case AssociatedGroupType.members:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedMemberGroup.Id);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedMemberGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedMemberGroup.Id.ToString();
                            }
                            break;
                        }
                    case AssociatedGroupType.visitors:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedVisitorGroup.Id);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedVisitorGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedVisitorGroup.Id.ToString();
                            }
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