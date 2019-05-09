using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;

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
            : base(web, GetGroupToken(groupType))
        {
            if (groupType == AssociatedGroupType.None)
            {
                throw new ArgumentOutOfRangeException(nameof(groupType));
            }
            _groupType = groupType;
        }

        internal AssociatedGroupType GroupType { get => _groupType; set => _groupType = value; }

        public override string GetReplaceValue()
        {

            if (string.IsNullOrEmpty(CacheValue))
            {
                switch (_groupType)
                {
                    case AssociatedGroupType.Owners:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedOwnerGroup.Title);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedOwnerGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedOwnerGroup.Title;
                            }
                            break;
                        }
                    case AssociatedGroupType.Members:
                        {
                            TokenContext.Load(TokenContext.Web, w => w.AssociatedMemberGroup.Title);
                            TokenContext.ExecuteQueryRetry();
                            if (!TokenContext.Web.AssociatedMemberGroup.ServerObjectIsNull.Value)
                            {
                                CacheValue = TokenContext.Web.AssociatedMemberGroup.Title;
                            }
                            break;
                        }
                    case AssociatedGroupType.Visitors:
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

        public static AssociatedGroupType GetGroupType(string token)
        {
            if (string.Equals(token, "{associatedownergroup}", StringComparison.OrdinalIgnoreCase))
            {
                return AssociatedGroupType.Owners;
            }

            if (string.Equals(token, "{associatedmembergroup}", StringComparison.OrdinalIgnoreCase))
            {
                return AssociatedGroupType.Members;
            }

            if (string.Equals(token, "{associatedvisitorgroup}", StringComparison.OrdinalIgnoreCase))
            {
                return AssociatedGroupType.Visitors;
            }

            return AssociatedGroupType.None;
        }

        public static string GetGroupToken(AssociatedGroupType groupType)
        {
            return $"{{associated{GetGroupTypeName(groupType)}group}}";
        }

        public static string GetGroupTypeName(AssociatedGroupType groupType)
        {
            switch (groupType)
            {
                case AssociatedGroupType.Owners:
                    return "owner";

                case AssociatedGroupType.Members:
                    return "member";

                case AssociatedGroupType.Visitors:
                    return "visitor";

                default:
                    throw new ArgumentOutOfRangeException(nameof(groupType));
            }
        }
    }
}