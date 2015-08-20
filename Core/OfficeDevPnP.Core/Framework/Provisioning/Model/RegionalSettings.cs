using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class RegionalSettings : IEquatable<RegionalSettings>
    {
        public Int32 AdjustHijriDays { get; set; }
        public Microsoft.SharePoint.Client.CalendarType AlternateCalendarType { get; set; }
        public Microsoft.SharePoint.Client.CalendarType CalendarType { get; set; }
        public Int32 Collation { get; set; }
        public DayOfWeek FirstDayOfWeek { get; set; }
        public Int32 FirstWeekOfYear { get; set; }
        public Int32 LocaleId { get; set; }
        public Boolean ShowWeeks { get; set; }
        public Boolean Time24 { get; set; }
        public Int32 TimeZone { get; set; }
        public WorkHour WorkDayEndHour { get; set; }
        public Int32 WorkDays { get; set; }
        public WorkHour WorkDayStartHour { get; set; }

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}",
                this.AdjustHijriDays,
                this.AlternateCalendarType,
                this.CalendarType,
                this.Collation,
                this.FirstDayOfWeek,
                this.FirstWeekOfYear,
                this.LocaleId,
                this.ShowWeeks,
                this.Time24,
                this.TimeZone,
                this.WorkDayEndHour,
                this.WorkDays,
                this.WorkDayStartHour
                ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is RegionalSettings))
            {
                return (false);
            }
            return (Equals((RegionalSettings)obj));
        }

        public bool Equals(RegionalSettings other)
        {
            return (this.AdjustHijriDays == other.AdjustHijriDays &&
                this.AlternateCalendarType == other.AlternateCalendarType &&
                this.CalendarType == other.CalendarType &&
                this.Collation == other.Collation &&
                this.FirstDayOfWeek == other.FirstDayOfWeek &&
                this.FirstWeekOfYear == other.FirstWeekOfYear &&
                this.LocaleId== other.LocaleId &&
                this.ShowWeeks == other.ShowWeeks &&
                this.Time24 == other.Time24 &&
                this.TimeZone == other.TimeZone &&
                this.WorkDayEndHour == other.WorkDayEndHour &&
                this.WorkDays == other.WorkDays &&
                this.WorkDayStartHour == other.WorkDayStartHour
                );
        }

        #endregion
    }

    public enum WorkHour
    {
        AM1200 = 0,
        AM0100 = 60,
        AM0200 = 120,
        AM0300 = 180,
        AM0400 = 240,
        AM0500 = 300,
        AM0600 = 360,
        AM0700 = 420,
        AM0800 = 480,
        AM0900 = 540,
        AM1000 = 600,
        AM1100 = 660,
        PM1200 = 720,
        PM0100 = 780,
        PM0200 = 840,
        PM0300 = 900,
        PM0400 = 960,
        PM0500 = 1020,
        PM0600 = 1080,
        PM0700 = 1140,
        PM0800 = 1200,
        PM0900 = 1260,
        PM1000 = 1320,
        PM1100 = 1380,
    }
}
