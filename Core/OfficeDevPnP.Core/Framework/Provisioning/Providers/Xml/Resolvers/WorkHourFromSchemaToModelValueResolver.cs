using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class WorkHourFromSchemaToModelValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        public object Resolve(object source, object destination, object sourceValue)
        {
            var workHour = sourceValue?.ToString();
            switch (workHour)
            {
                case "Item100AM":
                    return Model.WorkHour.AM0100;
                case "Item200AM":
                    return Model.WorkHour.AM0200;
                case "Item300AM":
                    return Model.WorkHour.AM0300;
                case "Item400AM":
                    return Model.WorkHour.AM0400;
                case "Item500AM":
                    return Model.WorkHour.AM0500;
                case "Item600AM":
                    return Model.WorkHour.AM0600;
                case "Item700AM":
                    return Model.WorkHour.AM0700;
                case "Item800AM":
                    return Model.WorkHour.AM0800;
                case "Item900AM":
                    return Model.WorkHour.AM0900;
                case "Item1000AM":
                    return Model.WorkHour.AM1000;
                case "Item1100AM":
                    return Model.WorkHour.AM1100;
                case "Item1200AM":
                    return Model.WorkHour.AM1200;
                case "Item100PM":
                    return Model.WorkHour.PM0100;
                case "Item200PM":
                    return Model.WorkHour.PM0200;
                case "Item300PM":
                    return Model.WorkHour.PM0300;
                case "Item400PM":
                    return Model.WorkHour.PM0400;
                case "Item500PM":
                    return Model.WorkHour.PM0500;
                case "Item600PM":
                    return Model.WorkHour.PM0600;
                case "Item700PM":
                    return Model.WorkHour.PM0700;
                case "Item800PM":
                    return Model.WorkHour.PM0800;
                case "Item900PM":
                    return Model.WorkHour.PM0900;
                case "Item1000PM":
                    return Model.WorkHour.PM1000;
                case "Item1100PM":
                    return Model.WorkHour.PM1100;
                case "Item1200PM":
                    return Model.WorkHour.PM1200;
                default:
                    return Model.WorkHour.AM0100;
            }
        }
    }
}
