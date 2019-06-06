using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{associatedownergroup}",
       Description = "Returns the title of the associated owners SharePoint group of a site",
       Example = "{associatedownersgroup}",
       Returns = "My Site Owners Group Title")]
    [TokenDefinitionDescription(
       Token = "{associatedmembergroup}",
       Description = "Returns the title of the associated members SharePoint group of a site",
       Example = "{associatedmembergroup}",
       Returns = "My Site Members Group Title")]
    [TokenDefinitionDescription(
       Token = "{associatedvisitorgroup}",
       Description = "Returns the title of the associated visitors SharePoint group of a site",
       Example = "{associatedvisitorgroup}",
       Returns = "My Site Visitors Group Title")]
    internal class AssociatedGroupToken : VolatileTokenDefinition
    {
        private AssociatedGroupType _groupType;

        public AssociatedGroupToken(Web web, AssociatedGroupType groupType)
            : base(web, $"{{associated{groupType.ToString().TrimEnd('s')}group}}")
        {
            _groupType = groupType;
        }

        internal AssociatedGroupType GroupType { get => _groupType; set => _groupType = value; }

        public override string GetReplaceValue()
        {

            if (string.IsNullOrEmpty(CacheValue))
            {
                switch (_groupType)
                {
                    case AssociatedGroupType.owners:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedOwnerGroup.Title);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedOwnerGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedOwnerGroup.Title;
                            }
                            break;
                        }
                    case AssociatedGroupType.members:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedMemberGroup.Title);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedMemberGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedMemberGroup.Title;
                            }
                            break;
                        }
                    case AssociatedGroupType.visitors:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedVisitorGroup.Title);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedVisitorGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedVisitorGroup.Title;
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