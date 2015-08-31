using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class AssociatedGroupToken : TokenDefinition
    {
        private AssociatedGroupType _groupType;

        public AssociatedGroupToken(Web web, AssociatedGroupType groupType)
            : base(web, string.Format("{{associated{0}group}}", groupType.ToString().TrimEnd('s')))
        {
            _groupType = groupType;
        }

        public override string GetReplaceValue()
        {

            if (string.IsNullOrEmpty(CacheValue))
            {
                var context = this.Web.Context as ClientContext;
                context.Load(Web, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                context.ExecuteQueryRetry();
                switch (_groupType)
                {
                    case AssociatedGroupType.owners:
                        {
                            CacheValue = Web.AssociatedOwnerGroup.Title;
                            break;
                        }
                    case AssociatedGroupType.members:
                        {
                            CacheValue = Web.AssociatedMemberGroup.Title;
                            break;
                        }
                    case AssociatedGroupType.visitors:
                        {
                            CacheValue = Web.AssociatedVisitorGroup.Title;
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