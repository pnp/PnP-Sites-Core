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
                web.Context.Load(web.RegionalSettings.TimeZone);
                web.Context.ExecuteQuery();

                Model.RegionalSettings settings = new Model.RegionalSettings();


                settings.AdjustHijriDays = web.RegionalSettings.AdjustHijriDays;
                settings.AlternateCalendarType = (CalendarType)web.RegionalSettings.AlternateCalendarType;
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
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.RegionalSettings.AdjustHijriDays = Convert.ToInt16(template.RegionalSettings.AdjustHijriDays);
                web.RegionalSettings.AlternateCalendarType = (short)template.RegionalSettings.AlternateCalendarType;
                web.RegionalSettings.Collation = Convert.ToInt16(template.RegionalSettings.Collation);
                web.RegionalSettings.FirstDayOfWeek = (uint)template.RegionalSettings.FirstDayOfWeek;
                web.RegionalSettings.FirstWeekOfYear = Convert.ToInt16(template.RegionalSettings.FirstWeekOfYear);
                web.RegionalSettings.LocaleId = Convert.ToUInt32(template.RegionalSettings.LocaleId);
                web.RegionalSettings.ShowWeeks = template.RegionalSettings.ShowWeeks;
                web.RegionalSettings.Time24 = template.RegionalSettings.Time24;
                web.RegionalSettings.TimeZone = web.RegionalSettings.TimeZones.GetById(template.RegionalSettings.TimeZone);
                web.RegionalSettings.WorkDayEndHour = (short)template.RegionalSettings.WorkDayEndHour;
                web.RegionalSettings.WorkDays = Convert.ToInt16(template.RegionalSettings.WorkDays);
                web.RegionalSettings.WorkDayStartHour = (short)template.RegionalSettings.WorkDayStartHour;
                web.RegionalSettings.Update();
                web.Context.ExecuteQueryRetry();
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return !web.IsSubSite();
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return !web.IsSubSite();
        }
    }
}
