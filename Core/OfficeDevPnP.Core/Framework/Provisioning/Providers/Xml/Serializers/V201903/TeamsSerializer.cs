using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Teams settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201903,
        SerializationSequence = -1, DeserializationSequence = -1,
        Default = false)]
    internal class TeamsSerializer : PnPBaseSchemaSerializer<ProvisioningTeams>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var teams = persistence.GetPublicInstancePropertyValue("Teams");

            if (teams != null)
            {
                var expressions = new Dictionary<Expression<Func<ProvisioningTeams, Object>>, IResolver>();

                // Manage Team Templates
                expressions.Add(t => t.TeamTemplates, new TeamTemplatesFromSchemaToModelTypeResolver());
                expressions.Add(t => t.TeamTemplates[0].JsonTemplate, new ExpressionValueResolver((s, v) => {
                    // Concatenate all the string values in the Text array of strings and return as the content of the JSON template
                    return ((s.GetPublicInstancePropertyValue("Text") as String[])?.Aggregate(String.Empty, (acc, next) => acc += (next != null ? next : String.Empty)));
                }));

                // Manage Teams
                expressions.Add(t => t.Teams, new TeamsFromSchemaToModelTypeResolver());
                expressions.Add(t => t.Teams[0].FunSettings, 
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamFunSettings>("FunSettings"));
                expressions.Add(t => t.Teams[0].GuestSettings, 
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamGuestSettings>("GuestSettings"));
                expressions.Add(t => t.Teams[0].MemberSettings,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamMemberSettings>("MembersSettings"));
                expressions.Add(t => t.Teams[0].MessagingSettings,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamMessagingSettings>("MessagingSettings"));
                expressions.Add(t => t.Teams[0].Security, new TeamSecurityFromSchemaToModelTypeResolver());

                expressions.Add(t => t.Teams[0].Channels[0].Tabs[0].Configuration,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamTabConfiguration>("Configuration"));

                // Handle the JSON content of the Message to send to the channel
                expressions.Add(t => t.Teams[0].Channels[0].Messages[0].Message, new ExpressionValueResolver((s, v) => s));

                PnPObjectsMapper.MapProperties(teams, template.ParentHierarchy.Teams, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ParentHierarchy != null && template.ParentHierarchy.Teams != null &&
                (template.ParentHierarchy.Teams.Apps != null ||
                template.ParentHierarchy.Teams.Teams != null ||
                template.ParentHierarchy.Teams.TeamTemplates != null))
            {
                var teamsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Teams, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var teamsType = Type.GetType(teamsTypeName, false);

                if (teamsType != null)
                {
                    var target = Activator.CreateInstance(teamsType, true);

                    var resolvers = new Dictionary<String, IResolver>();

                    //resolvers.Add($"{teamsType}.AppCatalog",
                    //    new AppCatalogFromModelToSchemaTypeResolver());
                    //resolvers.Add($"{teamsType}.ContentDeliveryNetwork",
                    //    new CdnFromModelToSchemaTypeResolver());
                    //resolvers.Add($"{teamsType}.SiteScripts",
                    //    new SiteScriptRefFromModelToSchemaTypeResolver());

                    //if (themeType != null)
                    //{
                    //    resolvers.Add($"{themeType}.Text",
                    //        new ExpressionValueResolver((s, v) =>
                    //        {
                    //            return (new String[] { (String)s.GetPublicInstancePropertyValue("Palette") });
                    //        }));
                    //}


                    PnPObjectsMapper.MapProperties(template.ParentHierarchy.Teams, target, resolvers, recursive: true);

                    if (target != null &&
                        (target.GetPublicInstancePropertyValue("Apps") != null ||
                        target.GetPublicInstancePropertyValue("Items") != null))
                    {
                        persistence.GetPublicInstanceProperty("Teams").SetValue(persistence, target);
                    }
                }
            }
        }
    }
}
