using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class CalendarTypeFromModelTypeResolver : IMemberValueResolver<object, object, Microsoft.SharePoint.Client.CalendarType, V201605.CalendarType>
    {

        public CalendarType Resolve(object source, object destination, Microsoft.SharePoint.Client.CalendarType sourceMember, CalendarType destMember, ResolutionContext context)
        {
            switch (sourceMember)
            {
                case Microsoft.SharePoint.Client.CalendarType.ChineseLunar:
                    return V201605.CalendarType.ChineseLunar;
                case Microsoft.SharePoint.Client.CalendarType.Gregorian:
                    return V201605.CalendarType.Gregorian;
                case Microsoft.SharePoint.Client.CalendarType.GregorianArabic:
                    return V201605.CalendarType.GregorianArabicCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianMEFrench:
                    return V201605.CalendarType.GregorianMiddleEastFrenchCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianXLITEnglish:
                    return V201605.CalendarType.GregorianTransliteratedEnglishCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianXLITFrench:
                    return V201605.CalendarType.GregorianTransliteratedFrenchCalendar;
                case Microsoft.SharePoint.Client.CalendarType.Hebrew:
                    return V201605.CalendarType.Hebrew;
                case Microsoft.SharePoint.Client.CalendarType.Hijri:
                    return V201605.CalendarType.Hijri;
                case Microsoft.SharePoint.Client.CalendarType.Japan:
                    return V201605.CalendarType.Japan;
                case Microsoft.SharePoint.Client.CalendarType.Korea:
                    return V201605.CalendarType.Korea;
                case Microsoft.SharePoint.Client.CalendarType.KoreaJapanLunar:
                    return V201605.CalendarType.KoreaandJapaneseLunar;
                case Microsoft.SharePoint.Client.CalendarType.SakaEra:
                    return V201605.CalendarType.SakaEra;
                case Microsoft.SharePoint.Client.CalendarType.Taiwan:
                    return V201605.CalendarType.Taiwan;
                case Microsoft.SharePoint.Client.CalendarType.Thai:
                    return V201605.CalendarType.Thai;
                case Microsoft.SharePoint.Client.CalendarType.UmAlQura:
                    return V201605.CalendarType.UmmalQura;
                default:
                    return V201605.CalendarType.None;
            }
        }
    }
}
