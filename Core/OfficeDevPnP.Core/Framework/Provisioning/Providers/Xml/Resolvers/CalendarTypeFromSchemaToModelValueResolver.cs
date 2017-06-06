using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class CalendarTypeFromSchemaToModelValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        public object Resolve(object source, object destination, object sourceValue)
        {
            var calendarType = sourceValue?.ToString();
            switch (calendarType)
            {
                case "ChineseLunar":
                    return Microsoft.SharePoint.Client.CalendarType.ChineseLunar;
                case "Gregorian":
                    return Microsoft.SharePoint.Client.CalendarType.Gregorian;
                case "GregorianArabicCalendar":
                    return Microsoft.SharePoint.Client.CalendarType.GregorianArabic;
                case "GregorianMiddleEastFrenchCalendar":
                    return Microsoft.SharePoint.Client.CalendarType.GregorianMEFrench;
                case "GregorianTransliteratedEnglishCalendar":
                    return Microsoft.SharePoint.Client.CalendarType.GregorianXLITEnglish;
                case "GregorianTransliteratedFrenchCalendar":
                    return Microsoft.SharePoint.Client.CalendarType.GregorianXLITFrench;
                case "Hebrew":
                    return Microsoft.SharePoint.Client.CalendarType.Hebrew;
                case "Hijri":
                    return Microsoft.SharePoint.Client.CalendarType.Hijri;
                case "Japan":
                    return Microsoft.SharePoint.Client.CalendarType.Japan;
                case "Korea":
                    return Microsoft.SharePoint.Client.CalendarType.Korea;
                case "KoreaandJapaneseLunar":
                    return Microsoft.SharePoint.Client.CalendarType.KoreaJapanLunar;
                case "SakaEra":
                    return Microsoft.SharePoint.Client.CalendarType.SakaEra;
                case "Taiwan":
                    return Microsoft.SharePoint.Client.CalendarType.Taiwan;
                case "Thai":
                    return Microsoft.SharePoint.Client.CalendarType.Thai;
                case "UmmalQura":
                    return Microsoft.SharePoint.Client.CalendarType.UmAlQura;
                case "None":
                default:
                    return Microsoft.SharePoint.Client.CalendarType.None;
            }
        }
    }
}
