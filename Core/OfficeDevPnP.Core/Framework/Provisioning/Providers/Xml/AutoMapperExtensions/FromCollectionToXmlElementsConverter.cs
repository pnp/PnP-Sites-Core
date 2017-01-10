using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class FromCollectionToXmlElementsConverter<TCollection, TAnyContainer> : ITypeConverter<TCollection, TAnyContainer>
    {
        public TAnyContainer Convert(TCollection source, TAnyContainer destination, ResolutionContext context)
        {
            var result = default(TAnyContainer);

            // TODO: Implement this converter

            return (result);
        }
    }
}
