using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for Teams from Schema to Model
    /// </summary>
    internal class TeamsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.Teams.Team>();

            var teams = source.GetPublicInstancePropertyValue("Items");
            var teamWithSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamWithSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamWithSettingsType = Type.GetType(teamWithSettingsTypeName, true);

            if (null != teams)
            {
                foreach (var t in ((IEnumerable)teams))
                {
                    if (teamWithSettingsType.IsInstanceOfType(t))
                    {
                        var targetItem = new Model.Teams.Team();
                        PnPObjectsMapper.MapProperties(t, targetItem, resolvers, recursive);
                        result.Add(targetItem);
                    }
                }
            }

            return (result);
        }
    }
}
