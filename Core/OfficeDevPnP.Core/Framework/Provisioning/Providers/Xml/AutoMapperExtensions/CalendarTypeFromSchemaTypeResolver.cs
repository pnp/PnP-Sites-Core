using AutoMapper;
using CalendarType = Microsoft.SharePoint.Client.CalendarType;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class CalendarTypeFromSchemaTypeResolver : IMemberValueResolver<object, object, V201605.CalendarType, CalendarType>
    {
        public CalendarType Resolve(object source, object destination, V201605.CalendarType sourceMember, CalendarType destMember, ResolutionContext context)
        {
            switch ((V201605.CalendarType)sourceMember)
            {
                case V201605.CalendarType.ChineseLunar:
                    return CalendarType.ChineseLunar;
                case V201605.CalendarType.Gregorian:
                    return CalendarType.Gregorian;
                case V201605.CalendarType.GregorianArabicCalendar:
                    return CalendarType.GregorianArabic;
                case V201605.CalendarType.GregorianMiddleEastFrenchCalendar:
                    return CalendarType.GregorianMEFrench;
                case V201605.CalendarType.GregorianTransliteratedEnglishCalendar:
                    return CalendarType.GregorianXLITEnglish;
                case V201605.CalendarType.GregorianTransliteratedFrenchCalendar:
                    return CalendarType.GregorianXLITFrench;
                case V201605.CalendarType.Hebrew:
                    return CalendarType.Hebrew;
                case V201605.CalendarType.Hijri:
                    return CalendarType.Hijri;
                case V201605.CalendarType.Japan:
                    return CalendarType.Japan;
                case V201605.CalendarType.Korea:
                    return CalendarType.Korea;
                case V201605.CalendarType.KoreaandJapaneseLunar:
                    return CalendarType.KoreaJapanLunar;
                case V201605.CalendarType.SakaEra:
                    return CalendarType.SakaEra;
                case V201605.CalendarType.Taiwan:
                    return CalendarType.Taiwan;
                case V201605.CalendarType.Thai:
                    return CalendarType.Thai;
                case V201605.CalendarType.UmmalQura:
                    return CalendarType.UmAlQura;
                default:
                    return CalendarType.None;
            }
        }

    }
}
