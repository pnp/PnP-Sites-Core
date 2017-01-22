using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class FromDecimalToDoubleResolver : IMemberValueResolver<object, object, decimal, double>
    {
        public double Resolve(object source, object destination, decimal sourceMember, double destMember, ResolutionContext context)
        {
            return ((double)sourceMember);
        }
    }
}
