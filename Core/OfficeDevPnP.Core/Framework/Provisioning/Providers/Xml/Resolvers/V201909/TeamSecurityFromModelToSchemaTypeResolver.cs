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
    /// Type resolver for Teams Security from Model to Schema
    /// </summary>
    internal class TeamSecurityFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Object result = null;

            var security = (source as Team)?.Security;
            if (null != security)
            {
                var teamSecurityTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSecurity, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var teamSecurityType = Type.GetType(teamSecurityTypeName, true);
                var teamSecurityUsersTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSecurityUsers, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var teamSecurityUsersType = Type.GetType(teamSecurityUsersTypeName, true);
                var teamSecurityUserTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSecurityUsersUser, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var teamSecurityUserType = Type.GetType(teamSecurityUserTypeName, true);

                result = Activator.CreateInstance(teamSecurityType);

                if (security.Owners != null && security.Owners.Count > 0)
                {
                    var owners = Activator.CreateInstance(teamSecurityUsersType);
                    owners.SetPublicInstancePropertyValue("ClearExistingItems", security.ClearExistingOwners);

                    var usersResolver = new CollectionFromModelToSchemaTypeResolver(teamSecurityUserType);
                    owners.SetPublicInstancePropertyValue("User", usersResolver.Resolve(security.Owners));

                    result.SetPublicInstancePropertyValue("Owners", owners);
                }

                if (security.Members != null && security.Members.Count > 0)
                {
                    var members = Activator.CreateInstance(teamSecurityUsersType);
                    members.SetPublicInstancePropertyValue("ClearExistingItems", security.ClearExistingOwners);

                    var usersResolver = new CollectionFromModelToSchemaTypeResolver(teamSecurityUserType);
                    members.SetPublicInstancePropertyValue("User", usersResolver.Resolve(security.Members));

                    result.SetPublicInstancePropertyValue("Members", members);
                }

                result.SetPublicInstancePropertyValue("AllowToAddGuests", security.AllowToAddGuests);
            }

            return (result);
        }
    }
}
