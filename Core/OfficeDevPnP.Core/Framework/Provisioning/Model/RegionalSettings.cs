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
        /// <summary>
        /// Constructor for RegionalSettings class
        /// </summary>
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
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
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

        /// <summary>
        /// Compares object with RegionalSettings
        /// </summary>
        /// <param name="obj">Object that represents RegionalSettings</param>
        /// <returns>true if the current object is equal to the RegionalSettings</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is RegionalSettings))
            {
                return (false);
            }
            return (Equals((RegionalSettings)obj));
        }

        /// <summary>
        /// Compares RegionalSettings object object based on AdjustHijriDays, AlternateCalendarType, CalendarType, Collation, FirstDayOfWeek, FirstWeekOfYear, LocaleId,
        /// ShowWeeks, Time24, TimeZone, WorkDayEndHour, WorkDays and WorkDayStartHour properties.
        /// </summary>
        /// <param name="other">RegionalSettings object</param>
        /// <returns>true if the RegionalSettings object is equal to the current object; otherwise, false.</returns>
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
        /// <summary>
        /// 00:00 am
        /// </summary>
        AM1200 = 0,
        /// <summary>
        /// 01:00 am
        /// </summary>
        AM0100 = 60,
        /// <summary>
        /// 02:00 am
        /// </summary>
        AM0200 = 120,
        /// <summary>
        /// 03:00 am
        /// </summary>
        AM0300 = 180,
        /// <summary>
        /// 04:00 am
        /// </summary>
        AM0400 = 240,
        /// <summary>
        /// 05:00 am
        /// </summary>
        AM0500 = 300,
        /// <summary>
        /// 06:00 am
        /// </summary>
        AM0600 = 360,
        /// <summary>
        /// 07:00 am
        /// </summary>
        AM0700 = 420,
        /// <summary>
        /// 08:00 am
        /// </summary>
        AM0800 = 480,
        /// <summary>
        /// 09:00 am
        /// </summary>
        AM0900 = 540,
        /// <summary>
        /// 10:00 am
        /// </summary>
        AM1000 = 600,
        /// <summary>
        /// 11:00 am
        /// </summary>
        AM1100 = 660,
        /// <summary>
        /// 12:00 pm
        /// </summary>
        PM1200 = 720,
        /// <summary>
        /// 01:00 pm
        /// </summary>
        PM0100 = 780,
        /// <summary>
        /// 02:00 pm
        /// </summary>
        PM0200 = 840,
        /// <summary>
        /// 03:00 pm
        /// </summary>
        PM0300 = 900,
        /// <summary>
        /// 04:00 pm
        /// </summary>
        PM0400 = 960,
        /// <summary>
        /// 05:00 pm
        /// </summary>
        PM0500 = 1020,
        /// <summary>
        /// 06:00 pm
        /// </summary>
        PM0600 = 1080,
        /// <summary>
        /// 07:00 pm
        /// </summary>
        PM0700 = 1140,
        /// <summary>
        /// 08:00 pm
        /// </summary>
        PM0800 = 1200,
        /// <summary>
        /// 09:00 pm
        /// </summary>
        PM0900 = 1260,
        /// <summary>
        /// 10:00 pm
        /// </summary>
        PM1000 = 1320,
        /// <summary>
        /// 11:00 pm
        /// </summary>
        PM1100 = 1380,
    }
}
