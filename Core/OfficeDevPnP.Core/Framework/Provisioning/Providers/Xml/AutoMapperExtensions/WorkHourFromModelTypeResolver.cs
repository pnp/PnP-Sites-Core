using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class WorkHourFromModelTypeResolver : IMemberValueResolver<object, object, Model.WorkHour, V201605.WorkHour>
    {
        public WorkHour Resolve(object source, object destination, Model.WorkHour sourceMember, WorkHour destMember, ResolutionContext context)
        {
            switch (sourceMember)
            {
                case Model.WorkHour.AM0100:
                    return V201605.WorkHour.Item100AM;
                case Model.WorkHour.AM0200:
                    return V201605.WorkHour.Item200AM;
                case Model.WorkHour.AM0300:
                    return V201605.WorkHour.Item300AM;
                case Model.WorkHour.AM0400:
                    return V201605.WorkHour.Item400AM;
                case Model.WorkHour.AM0500:
                    return V201605.WorkHour.Item500AM;
                case Model.WorkHour.AM0600:
                    return V201605.WorkHour.Item600AM;
                case Model.WorkHour.AM0700:
                    return V201605.WorkHour.Item700AM;
                case Model.WorkHour.AM0800:
                    return V201605.WorkHour.Item800AM;
                case Model.WorkHour.AM0900:
                    return V201605.WorkHour.Item900AM;
                case Model.WorkHour.AM1000:
                    return V201605.WorkHour.Item1000AM;
                case Model.WorkHour.AM1100:
                    return V201605.WorkHour.Item1100AM;
                case Model.WorkHour.AM1200:
                    return V201605.WorkHour.Item1200AM;
                case Model.WorkHour.PM0100:
                    return V201605.WorkHour.Item100PM;
                case Model.WorkHour.PM0200:
                    return V201605.WorkHour.Item200PM;
                case Model.WorkHour.PM0300:
                    return V201605.WorkHour.Item300PM;
                case Model.WorkHour.PM0400:
                    return V201605.WorkHour.Item400PM;
                case Model.WorkHour.PM0500:
                    return V201605.WorkHour.Item500PM;
                case Model.WorkHour.PM0600:
                    return V201605.WorkHour.Item600PM;
                case Model.WorkHour.PM0700:
                    return V201605.WorkHour.Item700PM;
                case Model.WorkHour.PM0800:
                    return V201605.WorkHour.Item800PM;
                case Model.WorkHour.PM0900:
                    return V201605.WorkHour.Item900PM;
                case Model.WorkHour.PM1000:
                    return V201605.WorkHour.Item1000PM;
                case Model.WorkHour.PM1100:
                    return V201605.WorkHour.Item1100PM;
                case Model.WorkHour.PM1200:
                    return V201605.WorkHour.Item1200PM;
                default:
                    return V201605.WorkHour.Item100AM;
            }
        }
    }
}
