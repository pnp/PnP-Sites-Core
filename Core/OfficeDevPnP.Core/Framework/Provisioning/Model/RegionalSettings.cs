using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Regional Settings for a site
    /// </summary>
    public partial class RegionalSettings : BaseModel, IEquatable<RegionalSettings>
    {

        public RegionalSettings()
        {
            AdjustHijriDays = 0;
            AlternateCalendarType = Microsoft.SharePoint.Client.CalendarType.None;
            CalendarType = Microsoft.SharePoint.Client.CalendarType.None;
            Collation = 0;
            FirstDayOfWeek = DayOfWeek.Sunday;
            FirstWeekOfYear = 0;
            LocaleId = 1033;
            ShowWeeks = false;
            WorkDayEndHour = WorkHour.PM0600;
            WorkDays = 5;
            WorkDayStartHour = WorkHour.AM0900;
        }

        #region Public Members

        /// <summary>
        /// The number of days to extend or reduce the current month in Hijri calendars
        /// </summary>
        public Int32 AdjustHijriDays { get; set; }

        /// <summary>
        /// The Alternate Calendar type that is used on the server
        /// </summary>
        public Microsoft.SharePoint.Client.CalendarType AlternateCalendarType { get; set; }

        /// <summary>
        /// The Calendar Type that is used on the server
        /// </summary>
        public Microsoft.SharePoint.Client.CalendarType CalendarType { get; set; }

        /// <summary>
        /// The Collation that is used on the site
        /// </summary>
        public Int32 Collation { get; set; }

        /// <summary>
        /// The First Day of the Week used in calendars on the server
        /// </summary>
        public DayOfWeek FirstDayOfWeek { get; set; }

        /// <summary>
        /// The First Week of the Year used in calendars on the server
        /// </summary>
        public Int32 FirstWeekOfYear { get; set; }

        /// <summary>
        /// The Locale Identifier in use on the server
        /// </summary>
        public Int32 LocaleId { get; set; }

        /// <summary>
        /// Defines whether to display the week number in day or week views of a calendar
        /// </summary>
        public Boolean ShowWeeks { get; set; }

        /// <summary>
        /// Defines whether to use a 24-hour time format in representing the hours of the day
        /// </summary>
        public Boolean Time24 { get; set; }

        /// <summary>
        /// The Time Zone that is used on the server
        /// </summary>
        public Int32 TimeZone { get; set; }

        /// <summary>
        /// The the default hour at which the work day ends on the calendar that is in use on the server
        /// </summary>
        public WorkHour WorkDayEndHour { get; set; }

        /// <summary>
        /// The work days of Web site calendars
        /// </summary>
        public Int32 WorkDays { get; set; }

        /// <summary>
        /// The the default hour at which the work day starts on the calendar that is in use on the server
        /// </summary>
        public WorkHour WorkDayStartHour { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|",
                this.AdjustHijriDays.GetHashCode(),
                this.AlternateCalendarType.GetHashCode(),
                this.CalendarType.GetHashCode(),
                this.Collation.GetHashCode(),
                this.FirstDayOfWeek.GetHashCode(),
                this.FirstWeekOfYear.GetHashCode(),
                this.LocaleId.GetHashCode(),
                this.ShowWeeks.GetHashCode(),
                this.Time24.GetHashCode(),
                this.TimeZone.GetHashCode(),
                this.WorkDayEndHour.GetHashCode(),
                this.WorkDays.GetHashCode(),
                this.WorkDayStartHour.GetHashCode()
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
            if (other == null)
            {
                return (false);
            }

            return (this.AdjustHijriDays == other.AdjustHijriDays &&
                this.AlternateCalendarType == other.AlternateCalendarType &&
                this.CalendarType == other.CalendarType &&
                this.Collation == other.Collation &&
                this.FirstDayOfWeek == other.FirstDayOfWeek &&
                this.FirstWeekOfYear == other.FirstWeekOfYear &&
                this.LocaleId == other.LocaleId &&
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

    /// <summary>
    /// The Work Hours of a Day
    /// </summary>
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
