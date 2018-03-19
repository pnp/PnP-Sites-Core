using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectRegionalSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Regional Settings"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {

                web.Context.Load(web.RegionalSettings);
                web.Context.Load(web.RegionalSettings.TimeZone, tz => tz.Id);
                web.Context.ExecuteQueryRetry();

                Model.RegionalSettings settings = new Model.RegionalSettings();

                settings.AdjustHijriDays = web.RegionalSettings.AdjustHijriDays;
                settings.AlternateCalendarType = (CalendarType)web.RegionalSettings.AlternateCalendarType;
                settings.CalendarType = (CalendarType)web.RegionalSettings.CalendarType;
                settings.Collation = web.RegionalSettings.Collation;
                settings.FirstDayOfWeek = (DayOfWeek)web.RegionalSettings.FirstDayOfWeek;
                settings.FirstWeekOfYear = web.RegionalSettings.FirstWeekOfYear;
                settings.LocaleId = (int)web.RegionalSettings.LocaleId;
                settings.ShowWeeks = web.RegionalSettings.ShowWeeks;
                settings.Time24 = web.RegionalSettings.Time24;
                settings.TimeZone = web.RegionalSettings.TimeZone.Id;
                settings.WorkDayEndHour = (WorkHour)web.RegionalSettings.WorkDayEndHour;
                settings.WorkDays = web.RegionalSettings.WorkDays;
                settings.WorkDayStartHour = (WorkHour)web.RegionalSettings.WorkDayStartHour;

                template.RegionalSettings = settings;

                // We're not comparing regional settings with the value stored in the base template as base templates are always for the US locale (1033)
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.Context.Load(web.RegionalSettings);
                web.Context.Load(web.RegionalSettings.TimeZone, tz => tz.Id);
                web.Context.ExecuteQueryRetry();

                var isDirty = false;
                if (web.RegionalSettings.AdjustHijriDays != template.RegionalSettings.AdjustHijriDays)
                {
                    web.RegionalSettings.AdjustHijriDays = Convert.ToInt16(template.RegionalSettings.AdjustHijriDays);
                    isDirty = true;
                }
                if (web.RegionalSettings.AlternateCalendarType != (short)template.RegionalSettings.AlternateCalendarType)
                {
                    web.RegionalSettings.AlternateCalendarType = (short)template.RegionalSettings.AlternateCalendarType;
                    isDirty = true;
                }
                if (template.RegionalSettings.CalendarType != CalendarType.None)
                {
                    if (web.RegionalSettings.CalendarType != (short)template.RegionalSettings.CalendarType)
                    {
                        web.RegionalSettings.CalendarType = (short)template.RegionalSettings.CalendarType;
                        isDirty = true;
                    }
                }
                if (web.RegionalSettings.Collation != Convert.ToInt16(template.RegionalSettings.Collation))
                {
                    web.RegionalSettings.Collation = Convert.ToInt16(template.RegionalSettings.Collation);
                    isDirty = true;
                }
                if (web.RegionalSettings.FirstDayOfWeek != (uint)template.RegionalSettings.FirstDayOfWeek)
                {
                    web.RegionalSettings.FirstDayOfWeek = (uint)template.RegionalSettings.FirstDayOfWeek;
                    isDirty = true;
                }
                if (web.RegionalSettings.FirstWeekOfYear != Convert.ToInt16(template.RegionalSettings.FirstWeekOfYear))
                {
                    web.RegionalSettings.FirstWeekOfYear = Convert.ToInt16(template.RegionalSettings.FirstWeekOfYear);
                    isDirty = true;
                }
                if (template.RegionalSettings.LocaleId > 0 && (web.RegionalSettings.LocaleId != Convert.ToUInt32(template.RegionalSettings.LocaleId)))
                {
                    web.RegionalSettings.LocaleId = Convert.ToUInt32(template.RegionalSettings.LocaleId);
                    isDirty = true;
                }
                if (web.RegionalSettings.ShowWeeks != template.RegionalSettings.ShowWeeks)
                {
                    web.RegionalSettings.ShowWeeks = template.RegionalSettings.ShowWeeks;
                    isDirty = true;
                }
                if (web.RegionalSettings.Time24 != template.RegionalSettings.Time24)
                {
                    web.RegionalSettings.Time24 = template.RegionalSettings.Time24;
                    isDirty = true;
                }
                if (template.RegionalSettings.TimeZone != 0 && (web.RegionalSettings.TimeZone.Id != template.RegionalSettings.TimeZone))
                {
                    web.RegionalSettings.TimeZone = web.RegionalSettings.TimeZones.GetById(template.RegionalSettings.TimeZone);
                    isDirty = true;
                }
                if (web.RegionalSettings.WorkDayEndHour != (short)template.RegionalSettings.WorkDayEndHour)
                {
                    web.RegionalSettings.WorkDayEndHour = (short)template.RegionalSettings.WorkDayEndHour;
                    isDirty = true;
                }
                if (template.RegionalSettings.WorkDays > 0 && (web.RegionalSettings.WorkDays != Convert.ToInt16(template.RegionalSettings.WorkDays)))
                {
                    web.RegionalSettings.WorkDays = Convert.ToInt16(template.RegionalSettings.WorkDays);
                    isDirty = true;
                }
                if (web.RegionalSettings.WorkDayStartHour != (short)template.RegionalSettings.WorkDayStartHour)
                {
                    web.RegionalSettings.WorkDayStartHour = (short)template.RegionalSettings.WorkDayStartHour;
                    isDirty = true;
                }
                if (isDirty)
                {
                    web.RegionalSettings.Update();
                    web.Context.ExecuteQueryRetry();
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return template.RegionalSettings != null;
        }
    }
}
