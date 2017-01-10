using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class FromListToTypedArrayResolver<TSource, TDestination, TDestMember> : IValueResolver<TSource, TDestination, TDestMember>
    {
        public TDestMember Resolve(TSource source, TDestination destination, TDestMember destMember, ResolutionContext context)
        {
            // TODO: Implement this resolver
            return (default(TDestMember));
        }
    }
}
