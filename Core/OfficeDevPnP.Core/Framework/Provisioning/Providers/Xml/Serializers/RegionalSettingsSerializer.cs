using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Regional Settings
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 400, DeserializationSequence = 400,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class RegionalSettingsSerializer : PnPBaseSchemaSerializer<Model.RegionalSettings>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var regionalSettings = persistence.GetPublicInstancePropertyValue("RegionalSettings");

            if (regionalSettings != null)
            {
                template.RegionalSettings = new Model.RegionalSettings();
                var expressions = new Dictionary<Expression<Func<Model.RegionalSettings, Object>>, IResolver>();

                expressions.Add(s => s.WorkDayStartHour, new WorkHourFromSchemaToModelValueResolver());
                expressions.Add(s => s.WorkDayEndHour, new WorkHourFromSchemaToModelValueResolver());
                expressions.Add(s => s.CalendarType, new CalendarTypeFromSchemaToModelValueResolver());
                expressions.Add(s => s.AlternateCalendarType, new CalendarTypeFromSchemaToModelValueResolver());
                expressions.Add(s => s.TimeZone, new ExpressionValueResolver((s, v) =>
                      !String.IsNullOrEmpty(v as string) ? Int32.Parse(v as string) : 0
                ));

                PnPObjectsMapper.MapProperties(regionalSettings, template.RegionalSettings, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.RegionalSettings != null)
            {
                var regionalSettingsType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.RegionalSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var target = Activator.CreateInstance(regionalSettingsType, true);
                var expressions = new Dictionary<string, IResolver>();
                expressions.Add($"{regionalSettingsType}.AdjustHijriDaysSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.AlternateCalendarTypeSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.CalendarTypeSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.CollationSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.FirstDayOfWeekSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.FirstWeekOfYearSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.LocaleIdSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.ShowWeeksSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.Time24Specified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.WorkDayEndHourSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.WorkDaysSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.WorkDayStartHourSpecified", new ExpressionValueResolver((s, p) => true));
                expressions.Add($"{regionalSettingsType}.WorkDayStartHour", new ExpressionValueResolver<Model.WorkHour>(v => v.FromTemplateToSchemaWorkHourV201605()));
                expressions.Add($"{regionalSettingsType}.WorkDayEndHour", new ExpressionValueResolver<Model.WorkHour>(v => v.FromTemplateToSchemaWorkHourV201605()));
                expressions.Add($"{regionalSettingsType}.CalendarType", new ExpressionValueResolver<CalendarType>(v => v.FromTemplateToSchemaCalendarTypeV201605()));
                expressions.Add($"{regionalSettingsType}.AlternateCalendarType", new ExpressionValueResolver<CalendarType>(v => v.FromTemplateToSchemaCalendarTypeV201605()));


                PnPObjectsMapper.MapProperties(template.RegionalSettings, target, expressions, recursive: true);

                persistence.GetPublicInstanceProperty("RegionalSettings").SetValue(persistence, target);
            }
        }
    }
}
