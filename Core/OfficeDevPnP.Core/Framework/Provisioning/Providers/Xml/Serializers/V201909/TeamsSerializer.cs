using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using ResolversV201909 = OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201909;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using OfficeDevPnP.Core.Extensions;
using System.Web;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers.V201909
{
    /// <summary>
    /// Class to serialize/deserialize the Teams settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201909,
        SerializationSequence = 100, DeserializationSequence = 100,
        Scope = SerializerScope.Tenant)]
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
                expressions.Add(t => t.TeamTemplates[0].JsonTemplate, new ExpressionValueResolver((s, v) =>
                {
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
                expressions.Add(t => t.Teams[0].DiscoverySettings,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamDiscoverySettings>("DiscoverySettings"));
                expressions.Add(t => t.Teams[0].Security,
                    new ResolversV201909.TeamSecurityFromSchemaToModelTypeResolver());

                expressions.Add(t => t.Teams[0].Channels[0].Tabs[0].Configuration,
                    new ComplexTypeFromSchemaToModelTypeResolver<TeamTabConfiguration>("Configuration"));

                // Handle the JSON content of the Message to send to the channel
                expressions.Add(t => t.Teams[0].Channels[0].Messages[0].Message, new ExpressionValueResolver((s, v) =>
                {
                    var message = s as string;
                    return HttpUtility.HtmlDecode(message);
                }));

                PnPObjectsMapper.MapProperties(teams, template.ParentHierarchy.Teams, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ParentHierarchy?.Teams != null &&
                (template.ParentHierarchy?.Teams?.Apps != null ||
                template.ParentHierarchy?.Teams?.Teams != null ||
                template.ParentHierarchy?.Teams?.TeamTemplates != null))
            {
                var teamsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Teams, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var teamsType = Type.GetType(teamsTypeName, false);

                if (teamsType != null)
                {
                    var teamTemplateTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamTemplate, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamTemplateType = Type.GetType(teamTemplateTypeName, true);
                    var teamWithSettingTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamWithSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamWithSettingType = Type.GetType(teamWithSettingTypeName, true);
                    var teamChannelTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamChannel, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamChannelType = Type.GetType(teamChannelTypeName, true);

                    var teamFunSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamWithSettingsFunSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamFunSettingsType = Type.GetType(teamFunSettingsTypeName, true);
                    var teamGuestSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamWithSettingsGuestSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamGuestSettingsType = Type.GetType(teamGuestSettingsTypeName, true);
                    var teamMembersSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamWithSettingsMembersSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamMembersSettingsType = Type.GetType(teamMembersSettingsTypeName, true);
                    var teamMessagingSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamWithSettingsMessagingSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamMessagingSettingsType = Type.GetType(teamMessagingSettingsTypeName, true);
                    var teamDiscoverySettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamWithSettingsDiscoverySettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamDiscoverySettingsType = Type.GetType(teamDiscoverySettingsTypeName, true);
                    var teamChannelTabTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamChannelTabsTab, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamChannelTabType = Type.GetType(teamChannelTabTypeName, true);
                    var teamChannelTabConfigurationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamChannelTabsTabConfiguration, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var teamChannelTabConfigurationType = Type.GetType(teamChannelTabConfigurationTypeName, true);

                    var target = Activator.CreateInstance(teamsType, true);

                    var resolvers = new Dictionary<String, IResolver>();

                    // Handle generic team objects (TeamTemplate and TeamWithSettings)
                    resolvers.Add($"{teamsType}.Items",
                        new TeamsItemsFromModelToSchemaTypeResolver());

                    // Handle JSON template for the TeamTemplate objects
                    resolvers.Add($"{teamTemplateType}.Text", new ExpressionValueResolver((s, v) =>
                    {
                        // Return the JSON template as text for the node content
                        return (new String[1] { (s as TeamTemplate)?.JsonTemplate });
                    }));

                    // Handle Security for TeamWithSettings
                    resolvers.Add($"{teamWithSettingType}.Security",
                        new ResolversV201909.TeamSecurityFromModelToSchemaTypeResolver());

                    // Handle all the settings for the TeamWithSettings objects
                    resolvers.Add($"{teamWithSettingType}.FunSettings",
                        new ComplexTypeFromModelToSchemaTypeResolver(teamFunSettingsType, "FunSettings"));
                    resolvers.Add($"{teamWithSettingType}.GuestSettings",
                        new ComplexTypeFromModelToSchemaTypeResolver(teamGuestSettingsType, "GuestSettings"));
                    resolvers.Add($"{teamWithSettingType}.MembersSettings",
                        new ComplexTypeFromModelToSchemaTypeResolver(teamMembersSettingsType, "MemberSettings"));
                    resolvers.Add($"{teamWithSettingType}.MessagingSettings",
                        new ComplexTypeFromModelToSchemaTypeResolver(teamMessagingSettingsType, "MessagingSettings"));
                    resolvers.Add($"{teamWithSettingType}.DiscoverySettings",
                        new ComplexTypeFromModelToSchemaTypeResolver(teamDiscoverySettingsType, "DiscoverySettings"));

                    // Handle channel Messages for TeamsWithSettings objects
                    resolvers.Add($"{teamChannelType}.Messages", new ExpressionValueResolver((s, v) =>
                    {
                        // Return the JSON messages as an array of Strings
                        return ((s as TeamChannel)?.Messages.Count > 0 ? (s as TeamChannel)?.Messages.Select(m => HttpUtility.HtmlEncode(m.Message)).ToArray() : null);
                    }));

                    // Handle channel Tab Configuration for TeamsWithSettings objects
                    resolvers.Add($"{teamChannelTabType}.Configuration",
                        new ComplexTypeFromModelToSchemaTypeResolver(teamChannelTabConfigurationType, "Configuration"));

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
