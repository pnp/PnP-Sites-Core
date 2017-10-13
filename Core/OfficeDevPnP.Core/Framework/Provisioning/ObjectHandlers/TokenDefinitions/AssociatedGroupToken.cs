using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;

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

                    int i = 0;
                    do
                    {

                        context.Load(context.Web, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title, w => w.HasUniqueRoleAssignments, w => w.Title);
                        context.ExecuteQueryRetry();

                        switch (_groupType)
                        {
                            case AssociatedGroupType.owners:
                            {
                                if (!context.Web.AssociatedOwnerGroup.ServerObjectIsNull())
                                    CacheValue = context.Web.AssociatedOwnerGroup.Title;
                                break;
                            }
                            case AssociatedGroupType.members:
                            {
                                if (!context.Web.AssociatedMemberGroup.ServerObjectIsNull())
                                    CacheValue = context.Web.AssociatedMemberGroup.Title;
                                break;
                            }
                            case AssociatedGroupType.visitors:
                            {
                                if (!context.Web.AssociatedVisitorGroup.ServerObjectIsNull())
                                    CacheValue = context.Web.AssociatedVisitorGroup.Title;
                                break;
                            }
                        }

                        if (string.IsNullOrEmpty(CacheValue) && context.Web.HasUniqueRoleAssignments && i==0)
                        {
                            var web = context.Web;
                            List<UserEntity> adminList = web.GetAdministrators();

                            web.Context.Load(web.CurrentUser, u => u.LoginName);

                            web.Context.ExecuteQueryRetry();

                            string secondaryLoginName = web.CurrentUser.LoginName;

                            web.CreateDefaultAssociatedGroups(
                                (adminList != null && adminList.Count > 0)
                                    ? adminList[0].LoginName
                                    : secondaryLoginName, secondaryLoginName, web.Title);

                            web.Context.ExecuteQueryRetry();
                        }
                        i++;
                    } while (string.IsNullOrEmpty(CacheValue) && i<2);
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