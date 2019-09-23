using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201909
{
    /// <summary>
    /// Type resolver for Teams Security from Schema to Model
    /// </summary>
    internal class TeamSecurityFromSchemaToModelTypeResolver: ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            TeamSecurity result = null;

            var security = source.GetPublicInstancePropertyValue("Security");
            if (null != security)
            {
                result = new TeamSecurity();

                // Process settings
                var clearExistingOwnersValue = (security?.GetPublicInstancePropertyValue("Owners")?.GetPublicInstancePropertyValue("ClearExistingItems"));
                result.ClearExistingOwners = clearExistingOwnersValue != null ? (Boolean)clearExistingOwnersValue : false;
                var clearExistingMembersValue = (security?.GetPublicInstancePropertyValue("Members")?.GetPublicInstancePropertyValue("ClearExistingItems"));
                result.ClearExistingMembers = clearExistingMembersValue != null ? (Boolean)clearExistingMembersValue : false;

                // Process users (owners and members)
                var usersResolver = new CollectionFromSchemaToModelTypeResolver(typeof(TeamSecurityUser));

                var owners = security.GetPublicInstancePropertyValue("Owners");
                if (null != owners)
                {
                    result.Owners.AddRange(
                        usersResolver.Resolve(owners.GetPublicInstancePropertyValue("User")) 
                        as IEnumerable<TeamSecurityUser>);
                }

                var members = security.GetPublicInstancePropertyValue("Members");
                if (null != members)
                {
                    result.Members.AddRange(
                        usersResolver.Resolve(members.GetPublicInstancePropertyValue("User")) 
                        as IEnumerable<TeamSecurityUser>);
                }

                var allowToAddGuests = security.GetPublicInstancePropertyValue("AllowToAddGuests");
                result.AllowToAddGuests = allowToAddGuests != null ? bool.Parse(allowToAddGuests.ToString()) : false;
            }

            return (result);
        }
    }
}
