using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for TeamTemplates from Schema to Model
    /// </summary>
    internal class TeamTemplatesFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.Teams.TeamTemplate>();

            var teamTemplates = source.GetPublicInstancePropertyValue("Items");
            var teamTemplateTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamTemplate, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamTemplateType = Type.GetType(teamTemplateTypeName, true);

            resolvers = new Dictionary<string, IResolver>();
            resolvers.Add($"{typeof(Model.Teams.TeamTemplate).FullName}.JsonTemplate", new ExpressionValueResolver((s, v) => {
                // Concatenate all the string values in the Text array of strings and return as the content of the JSON template
                return (((String[])s.GetPublicInstancePropertyValue("Text")).Aggregate(String.Empty, (acc, next) => acc += (next != null ? next : String.Empty)));
            }));

            if (null != teamTemplates)
            {
                foreach (var t in ((IEnumerable)teamTemplates))
                {
                    if (teamTemplateType.IsInstanceOfType(t))
                    {
                        var targetItem = new Model.Teams.TeamTemplate();
                        PnPObjectsMapper.MapProperties(t, targetItem, resolvers, recursive);
                        result.Add(targetItem);
                    }
                }
            }

            return (result);
        }
    }
}
