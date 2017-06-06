using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
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
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    context.Load(context.Web, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                    context.ExecuteQueryRetry();
                    switch (_groupType)
                    {
                        case AssociatedGroupType.owners:
                            {
                                CacheValue = context.Web.AssociatedOwnerGroup.Title;
                                break;
                            }
                        case AssociatedGroupType.members:
                            {
                                CacheValue = context.Web.AssociatedMemberGroup.Title;
                                break;
                            }
                        case AssociatedGroupType.visitors:
                            {
                                CacheValue = context.Web.AssociatedVisitorGroup.Title;
                                break;
                            }
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