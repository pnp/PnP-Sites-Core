using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class WorkHourFromSchemaTypeResolver : IMemberValueResolver<object, object, V201605.WorkHour, Model.WorkHour>
    {
        public Model.WorkHour Resolve(object source, object destination, V201605.WorkHour sourceMember, Model.WorkHour destMember, ResolutionContext context)
        {
            switch (sourceMember)
            {
                case V201605.WorkHour.Item100AM:
                    return Model.WorkHour.AM0100;
                case V201605.WorkHour.Item200AM:
                    return Model.WorkHour.AM0200;
                case V201605.WorkHour.Item300AM:
                    return Model.WorkHour.AM0300;
                case V201605.WorkHour.Item400AM:
                    return Model.WorkHour.AM0400;
                case V201605.WorkHour.Item500AM:
                    return Model.WorkHour.AM0500;
                case V201605.WorkHour.Item600AM:
                    return Model.WorkHour.AM0600;
                case V201605.WorkHour.Item700AM:
                    return Model.WorkHour.AM0700;
                case V201605.WorkHour.Item800AM:
                    return Model.WorkHour.AM0800;
                case V201605.WorkHour.Item900AM:
                    return Model.WorkHour.AM0900;
                case V201605.WorkHour.Item1000AM:
                    return Model.WorkHour.AM1000;
                case V201605.WorkHour.Item1100AM:
                    return Model.WorkHour.AM1100;
                case V201605.WorkHour.Item1200AM:
                    return Model.WorkHour.AM1200;
                case V201605.WorkHour.Item100PM:
                    return Model.WorkHour.PM0100;
                case V201605.WorkHour.Item200PM:
                    return Model.WorkHour.PM0200;
                case V201605.WorkHour.Item300PM:
                    return Model.WorkHour.PM0300;
                case V201605.WorkHour.Item400PM:
                    return Model.WorkHour.PM0400;
                case V201605.WorkHour.Item500PM:
                    return Model.WorkHour.PM0500;
                case V201605.WorkHour.Item600PM:
                    return Model.WorkHour.PM0600;
                case V201605.WorkHour.Item700PM:
                    return Model.WorkHour.PM0700;
                case V201605.WorkHour.Item800PM:
                    return Model.WorkHour.PM0800;
                case V201605.WorkHour.Item900PM:
                    return Model.WorkHour.PM0900;
                case V201605.WorkHour.Item1000PM:
                    return Model.WorkHour.PM1000;
                case V201605.WorkHour.Item1100PM:
                    return Model.WorkHour.PM1100;
                case V201605.WorkHour.Item1200PM:
                    return Model.WorkHour.PM1200;
                default:
                    return Model.WorkHour.AM0100;
            }
        }
    }
}
