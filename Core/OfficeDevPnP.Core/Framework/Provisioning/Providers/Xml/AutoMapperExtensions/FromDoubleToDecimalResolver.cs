using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class FromDoubleToDecimalResolver : IMemberValueResolver<object, object, double, decimal>
    {
        public decimal Resolve(object source, object destination, double sourceMember, decimal destMember, ResolutionContext context)
        {
            return ((decimal)sourceMember);
        }
    }
}
